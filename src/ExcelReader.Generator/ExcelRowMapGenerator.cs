using System.Diagnostics.CodeAnalysis;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

[assembly: SuppressMessage("Globalization", "CA1307:Specify StringComparison for clarity",
    Justification = "The string.Replace(string,string,StringComparison) overload doesn't exist on netstandard2.0, this project's required TFM (see the .csproj comment). The 2-arg overload is already ordinal.")]
namespace ExcelReader.Generator
{
    /// <summary>
    /// Emits <c>IExcelRowMap&lt;T&gt;</c>/<c>IExcelRecordMap&lt;T&gt;</c> implementations for every type
    /// marked <c>[ExcelSerializable]</c>, from the same <c>[ExcelColumn]</c>/<c>[ExcelRequired]</c>/
    /// <c>[ExcelIgnore]</c> attributes the reflection-based <c>TypeMapper&lt;T&gt;</c> reads — so typed
    /// reading and writing work under trimming/Native AOT without reflecting over the marked type at
    /// runtime.
    /// </summary>
    [Generator(LanguageNames.CSharp)]
    public sealed class ExcelRowMapGenerator : IIncrementalGenerator
    {
        private const string SerializableAttribute = "ExcelReader.Core.Parser.ExcelSerializableAttribute";
        private const string ColumnAttribute = "ExcelReader.Core.Parser.ExcelColumnAttribute";
        private const string RequiredAttribute = "ExcelReader.Core.Parser.ExcelRequiredAttribute";
        private const string IgnoreAttribute = "ExcelReader.Core.Parser.ExcelIgnoreAttribute";

        private static readonly DiagnosticDescriptor NotPartialDescriptor = new(
            "EXR001",
            "Type marked [ExcelSerializable] must be declared partial",
            "Type '{0}' is marked [ExcelSerializable] but is not declared 'partial'; add 'partial' to its declaration",
            "ExcelReader.Generator",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        private static readonly DiagnosticDescriptor ContainingTypeNotPartialDescriptor = new(
            "EXR002",
            "Type containing an [ExcelSerializable] type must be declared partial",
            "Type '{0}' is nested inside '{1}', which is not declared 'partial'; add 'partial' to '{1}'",
            "ExcelReader.Generator",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        /// <inheritdoc/>
        public void Initialize(IncrementalGeneratorInitializationContext context)
        {
            IncrementalValuesProvider<INamedTypeSymbol> candidates = context.SyntaxProvider
                .ForAttributeWithMetadataName(
                    SerializableAttribute,
                    predicate: static (node, _) => node is ClassDeclarationSyntax or StructDeclarationSyntax,
                    transform: static (ctx, _) => (INamedTypeSymbol)ctx.TargetSymbol);

            context.RegisterSourceOutput(candidates, static (spc, symbol) => Emit(spc, symbol));
        }

        private static void Emit(SourceProductionContext context, INamedTypeSymbol symbol)
        {
            if (!IsPartial(symbol))
            {
                context.ReportDiagnostic(Diagnostic.Create(NotPartialDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name));
                return;
            }
            for (INamedTypeSymbol? outer = symbol.ContainingType; outer is not null; outer = outer.ContainingType)
            {
                if (!IsPartial(outer))
                {
                    context.ReportDiagnostic(Diagnostic.Create(ContainingTypeNotPartialDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name, outer.Name));
                    return;
                }
            }

            PropertyPlan[] properties = [.. symbol.GetMembers()
                                                  .OfType<IPropertySymbol>()
                                                  .Where(static p => !p.IsStatic && p.Parameters.Length == 0 && p.DeclaredAccessibility == Accessibility.Public && !HasAttribute(p, IgnoreAttribute))
                                                  .Select(BuildPlan)
                                                  .Where(static p => p.Read is not null || p.CanWrite)];

            string source = GenerateSource(symbol, properties);
            string hintName = symbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat)
                .Replace("global::", "")
                .Replace('.', '_');
            context.AddSource($"{hintName}.ExcelRowMap.g.cs", source);
        }

        private static bool IsPartial(INamedTypeSymbol symbol)
        {
            return symbol.DeclaringSyntaxReferences
                .Select(static r => r.GetSyntax())
                .OfType<TypeDeclarationSyntax>()
                .Any(static d => d.Modifiers.Any(SyntaxKind.PartialKeyword));
        }

        private static bool HasAttribute(IPropertySymbol property, string fullName)
        {
            return property.GetAttributes().Any(a => string.Equals(a.AttributeClass?.ToDisplayString(), fullName, StringComparison.Ordinal));
        }

        // A2 scope: only string and int properties are supported so far — everything else is silently
        // skipped (same "no parser, no binding" outcome the reflection path gives an unrecognized type),
        // expanded to the rest of ColumnParserFactory's type table in a later phase.
        private static PropertyPlan BuildPlan(IPropertySymbol property)
        {
            string[] names = [.. property.GetAttributes()
                .Where(a => string.Equals(a.AttributeClass?.ToDisplayString(), ColumnAttribute, StringComparison.Ordinal))
                .Select(static a => (string?)a.ConstructorArguments.FirstOrDefault().Value)
                .Where(static n => n is not null)!];
            if (names.Length == 0)
            {
                names = [property.Name];
            }

            AttributeData? required = property.GetAttributes()
                .FirstOrDefault(a => string.Equals(a.AttributeClass?.ToDisplayString(), RequiredAttribute, StringComparison.Ordinal));
            bool isRequired = required is not null;
            bool allowEmpty = required?.NamedArguments.FirstOrDefault(static kv => string.Equals(kv.Key, "AllowEmpty", StringComparison.Ordinal)).Value.Value is true;
            bool requireValue = isRequired && !allowEmpty;

            string? reader = property.Type.SpecialType switch
            {
                SpecialType.System_String => "global::ExcelReader.Core.Parser.ExcelCellReaders.String",
                SpecialType.System_Int32 => "global::ExcelReader.Core.Parser.ExcelCellReaders.Parsable<int>",
                _ => null,
            };

            bool canRead = reader is not null && property.SetMethod is { DeclaredAccessibility: Accessibility.Public };
            bool canWrite = (property.Type.SpecialType == SpecialType.System_String || property.Type.SpecialType == SpecialType.System_Int32)
                && property.GetMethod is { DeclaredAccessibility: Accessibility.Public };

            return new PropertyPlan(
                property.Name,
                names,
                canRead ? reader : null,
                canWrite,
                isRequired,
                requireValue);
        }

        private static string GenerateSource(INamedTypeSymbol symbol, PropertyPlan[] properties)
        {
            string qualifiedType = symbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
            var sb = new StringBuilder();
            sb.AppendLine("// <auto-generated/>");
            sb.AppendLine("#nullable enable");

            string? ns = symbol.ContainingNamespace.IsGlobalNamespace ? null : symbol.ContainingNamespace.ToDisplayString();
            if (ns is not null)
            {
                sb.AppendLine($"namespace {ns}");
                sb.AppendLine("{");
            }

            var containers = new List<INamedTypeSymbol>();
            for (INamedTypeSymbol? outer = symbol.ContainingType; outer is not null; outer = outer.ContainingType)
            {
                containers.Insert(0, outer);
            }
            foreach (INamedTypeSymbol outer in containers)
            {
                sb.AppendLine($"partial {(outer.TypeKind == TypeKind.Struct ? "struct" : "class")} {outer.Name}");
                sb.AppendLine("{");
            }

            string kind = symbol.TypeKind == TypeKind.Struct ? "struct" : "class";
            sb.AppendLine($"partial {kind} {symbol.Name} : global::ExcelReader.Core.Parser.IExcelRowMap<{qualifiedType}>, global::ExcelReader.Core.Writer.IExcelRecordMap<{qualifiedType}>");
            sb.AppendLine("{");

            AppendRowMap(sb, symbol, qualifiedType, properties);
            AppendRecordMap(sb, qualifiedType, properties);

            sb.AppendLine("}"); // type
            for (int i = 0; i < containers.Count; i++)
            {
                sb.AppendLine("}"); // containers
            }
            if (ns is not null)
            {
                sb.AppendLine("}"); // namespace
            }
            return sb.ToString();
        }

        private static void AppendRowMap(StringBuilder sb, INamedTypeSymbol symbol, string qualifiedType, PropertyPlan[] properties)
        {
            sb.AppendLine($"    public static void ConfigureExcelRowMap(global::ExcelReader.Core.Parser.ExcelRowMapBuilder<{qualifiedType}> builder)");
            sb.AppendLine("    {");
            bool useDefault = symbol.TypeKind == TypeKind.Struct
                && !symbol.Constructors.Any(static c => c.Parameters.Length == 0 && !c.IsImplicitlyDeclared);
            sb.AppendLine(useDefault
                ? "        builder.Factory(null)"
                : $"        builder.Factory(static () => new {qualifiedType}())");
            foreach (PropertyPlan p in properties)
            {
                if (p.Read is null)
                {
                    continue;
                }
                string names = string.Join(", ", p.HeaderNames.Select(static n => $"\"{n.Replace("\"", "\\\"")}\""));
                sb.AppendLine($"            .Property([{names}], {p.Read}, static (ref {qualifiedType} m, {PropertyValueType(p)} v) => m.{p.PropertyName} = v, isRequired: {Bool(p.IsRequired)}, requireValue: {Bool(p.RequireValue)})");
            }
            sb.AppendLine("        ;");
            sb.AppendLine("    }");
        }

        private static void AppendRecordMap(StringBuilder sb, string qualifiedType, PropertyPlan[] properties)
        {
            sb.AppendLine($"    public static void ConfigureExcelRecordMap(global::ExcelReader.Core.Writer.ExcelRecordMapBuilder<{qualifiedType}> builder)");
            sb.AppendLine("    {");
            sb.AppendLine("        builder");
            foreach (PropertyPlan p in properties)
            {
                if (!p.CanWrite)
                {
                    continue;
                }
                string header = p.HeaderNames[0].Replace("\"", "\\\"");
                sb.AppendLine($"            .Column(\"{header}\", static (row, m) => row.Write(m.{p.PropertyName}))");
            }
            sb.AppendLine("        ;");
            sb.AppendLine("    }");
        }

        // Only string/int are supported (A2 scope), so the setter's value type is always one of these two.
        private static string PropertyValueType(PropertyPlan p)
        {
            return string.Equals(p.Read, "global::ExcelReader.Core.Parser.ExcelCellReaders.String", StringComparison.Ordinal) ? "string" : "int";
        }

        private static string Bool(bool value)
        {
            return value ? "true" : "false";
        }

        // netstandard2.1 has no IsExternalInit, so record struct (init-only accessors) isn't available
        // here — a plain constructor-initialized struct does the same job.
        private readonly struct PropertyPlan
        {
            internal PropertyPlan(string propertyName, string[] headerNames, string? read, bool canWrite, bool isRequired, bool requireValue)
            {
                PropertyName = propertyName;
                HeaderNames = headerNames;
                Read = read;
                CanWrite = canWrite;
                IsRequired = isRequired;
                RequireValue = requireValue;
            }

            internal string PropertyName { get; }
            internal string[] HeaderNames { get; }
            internal string? Read { get; }
            internal bool CanWrite { get; }
            internal bool IsRequired { get; }
            internal bool RequireValue { get; }
        }
    }
}
