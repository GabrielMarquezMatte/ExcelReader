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
    /// <c>[ExcelConverter]</c>/<c>[ExcelIgnore]</c> attributes the reflection-based <c>TypeMapper&lt;T&gt;</c>
    /// reads — so typed reading and writing work under trimming/Native AOT without reflecting over the
    /// marked type at runtime.
    /// </summary>
    [Generator(LanguageNames.CSharp)]
    public sealed class ExcelRowMapGenerator : IIncrementalGenerator
    {
        private const string SerializableAttribute = "ExcelReader.Core.Parser.ExcelSerializableAttribute";
        private const string ColumnAttribute = "ExcelReader.Core.Parser.ExcelColumnAttribute";
        private const string RequiredAttribute = "ExcelReader.Core.Parser.ExcelRequiredAttribute";
        private const string IgnoreAttribute = "ExcelReader.Core.Parser.ExcelIgnoreAttribute";
        private const string ConverterAttribute = "ExcelReader.Core.Parser.ExcelConverterAttribute";
        private const string CellConverterInterface = "ExcelReader.Core.Parser.IExcelCellConverter<T>";
        private const string CellWriterInterface = "ExcelReader.Core.Writer.IExcelCellWriter<T>";

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

        private static readonly DiagnosticDescriptor RequiredWithNoParserDescriptor = new(
            "EXR003",
            "[ExcelRequired] property has no available reader",
            "Property '{0}.{1}' is marked [ExcelRequired] but its type '{2}' has no built-in reader; add an [ExcelConverter] for it",
            "ExcelReader.Generator",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        private static readonly DiagnosticDescriptor ConverterTypeMismatchDescriptor = new(
            "EXR004",
            "[ExcelConverter] type does not implement IExcelCellConverter<T> for the property's exact type",
            "Converter '{0}' must implement IExcelCellConverter<{1}> to convert property '{2}.{3}'",
            "ExcelReader.Generator",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        private static readonly DiagnosticDescriptor DuplicateHeaderDescriptor = new(
            "EXR006",
            "Two properties bind the same header name",
            "Header name '{0}' is bound by more than one property of '{1}'; only the first one encountered will ever match",
            "ExcelReader.Generator",
            DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        private enum ReadKind
        {
            None,
            Value,
            Nullable,
            GuidValue,
            GuidNullable,
            Converted,
        }

        private enum WriteKind
        {
            None,
            Direct,
            ToStringFallback,
        }

        // Matches RowWriteMethods<TRow>.Numeric exactly (the reflection write path's hashset) — only
        // these get the generic, numeric-cell Write<T> overload. Everything else that's still writable
        // (enum, Guid, char, Half, Int128, UInt128, TimeSpan, DateTimeOffset — every other
        // IUtf8SpanParsable/enum type this generator's read side supports) writes as text via ToString(),
        // even though several of them (Guid, TimeSpan, DateTimeOffset, char, Half, Int128, UInt128) also
        // implement IUtf8SpanFormattable and so *could* go through the generic path — using that path for
        // them anyway would write a different cell type than the reflection path does for the same model.
        private static readonly HashSet<SpecialType> NumericWriteSpecialTypes =
        [
            SpecialType.System_Byte, SpecialType.System_SByte, SpecialType.System_Int16, SpecialType.System_UInt16,
            SpecialType.System_Int32, SpecialType.System_UInt32, SpecialType.System_Int64, SpecialType.System_UInt64,
            SpecialType.System_Single, SpecialType.System_Double, SpecialType.System_Decimal,
        ];

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

            IPropertySymbol[] candidateProperties = [.. symbol.GetMembers()
                .OfType<IPropertySymbol>()
                .Where(static p => !p.IsStatic && p.Parameters.Length == 0 && p.DeclaredAccessibility == Accessibility.Public && !HasAttribute(p, IgnoreAttribute))];

            var boundHeaders = new HashSet<string>(StringComparer.Ordinal);
            var plans = new List<PropertyPlan>(candidateProperties.Length);
            foreach (IPropertySymbol property in candidateProperties)
            {
                PropertyPlan plan = BuildPlan(context, symbol, property);
                if (plan.Read.Kind != ReadKind.None)
                {
                    foreach (string name in plan.HeaderNames.Where(n => !boundHeaders.Add(n)))
                    {
                        context.ReportDiagnostic(Diagnostic.Create(DuplicateHeaderDescriptor, symbol.Locations.FirstOrDefault(), name, symbol.Name));
                    }
                }
                if (plan.Read.Kind != ReadKind.None || plan.WriteEmit is not null)
                {
                    plans.Add(plan);
                }
            }

            string source = GenerateSource(symbol, plans);
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

        private static PropertyPlan BuildPlan(SourceProductionContext context, INamedTypeSymbol owner, IPropertySymbol property)
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

            bool canSet = property.SetMethod is { DeclaredAccessibility: Accessibility.Public };
            bool canGet = property.GetMethod is { DeclaredAccessibility: Accessibility.Public };
            string qualifiedProperty = property.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);

            AttributeData? converterAttr = property.GetAttributes()
                .FirstOrDefault(a => string.Equals(a.AttributeClass?.ToDisplayString(), ConverterAttribute, StringComparison.Ordinal));

            ReadPlan read = default;
            string? writeEmit = null;
            string? converterFieldDecl = null;

            if (converterAttr is not null && converterAttr.ConstructorArguments.FirstOrDefault().Value is ITypeSymbol converterType)
            {
                string converterQualified = converterType.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                string fieldName = $"s_converter_{property.Name}";
                bool implementsConverter = ImplementsGenericInterface(converterType, CellConverterInterface, property.Type);
                bool implementsWriter = ImplementsGenericInterface(converterType, CellWriterInterface, property.Type);
                if (implementsConverter || implementsWriter)
                {
                    converterFieldDecl = $"    private static readonly {converterQualified} {fieldName} = new();";
                }
                if (implementsConverter && canSet)
                {
                    read = new ReadPlan(ReadKind.Converted, fieldName, qualifiedProperty);
                }
                else if (!implementsConverter)
                {
                    context.ReportDiagnostic(Diagnostic.Create(
                        ConverterTypeMismatchDescriptor, property.Locations.FirstOrDefault(), converterQualified, qualifiedProperty, owner.Name, property.Name));
                }
                if (implementsWriter && canGet)
                {
                    writeEmit = $"            .Column(\"{names[0].Replace("\"", "\\\"")}\", static (row, m) => {fieldName}.Write(row, m.{property.Name}))";
                }
            }
            else
            {
                bool isNullable = TryGetNullableUnderlying(property.Type, out ITypeSymbol underlying);
                (string Reader, string ValueType, bool IsGuid)? builtin = TryGetBuiltInReader(underlying);
                if (builtin is { } b && canSet)
                {
                    read = new ReadPlan(SelectReadKind(b.IsGuid, isNullable), b.Reader, b.ValueType);
                }
                else if (isRequired)
                {
                    context.ReportDiagnostic(Diagnostic.Create(
                        RequiredWithNoParserDescriptor, property.Locations.FirstOrDefault(), owner.Name, property.Name, qualifiedProperty));
                }

                WriteKind writeKind = GetWriteKind(underlying);
                if (canGet && writeKind != WriteKind.None)
                {
                    string valueExpr = WriteValueExpression(writeKind, isNullable, property.Name);
                    writeEmit = $"            .Column(\"{names[0].Replace("\"", "\\\"")}\", static (row, m) => row.Write({valueExpr}))";
                }
            }

            return new PropertyPlan(property.Name, names, read, isRequired, requireValue, writeEmit, converterFieldDecl);
        }

        private static ReadKind SelectReadKind(bool isGuid, bool isNullable)
        {
            if (isGuid)
            {
                return isNullable ? ReadKind.GuidNullable : ReadKind.GuidValue;
            }
            return isNullable ? ReadKind.Nullable : ReadKind.Value;
        }

        private static string WriteValueExpression(WriteKind kind, bool isNullable, string propertyName)
        {
            if (kind == WriteKind.Direct)
            {
                return $"m.{propertyName}";
            }
            return isNullable ? $"m.{propertyName}?.ToString()" : $"m.{propertyName}.ToString()";
        }

        private static WriteKind GetWriteKind(ITypeSymbol underlying)
        {
            if (underlying.SpecialType is SpecialType.System_String or SpecialType.System_Boolean
                || IsSystemType(underlying, "DateTime") || IsSystemType(underlying, "DateOnly") || IsSystemType(underlying, "TimeOnly")
                || NumericWriteSpecialTypes.Contains(underlying.SpecialType))
            {
                return WriteKind.Direct;
            }
            if (underlying.TypeKind == TypeKind.Enum || IsSystemType(underlying, "Guid") || IsUtf8SpanParsable(underlying))
            {
                return WriteKind.ToStringFallback;
            }
            return WriteKind.None;
        }

        // Guid implements IUtf8SpanParsable<Guid> starting only on the newer TFM ColumnParserFactory's
        // own conditional compilation gates on; the generated code targets whatever TFM the *consumer*
        // builds for, so both branches are emitted guarded the same way, and the C# compiler picks the
        // live one per consumer TFM — mirroring ColumnParserFactory.ReadGuid/BuildParsableCore inside
        // ExcelReader.Core itself.
        private static (string Reader, string ValueType, bool IsGuid)? TryGetBuiltInReader(ITypeSymbol underlying)
        {
            if (underlying.SpecialType == SpecialType.System_String)
            {
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.String", "string", false);
            }
            if (underlying.SpecialType == SpecialType.System_Boolean)
            {
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.Bool", "bool", false);
            }
            if (IsSystemType(underlying, "DateTime"))
            {
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.DateTimeSerial", "global::System.DateTime", false);
            }
            if (IsSystemType(underlying, "DateOnly"))
            {
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.DateOnlySerial", "global::System.DateOnly", false);
            }
            if (IsSystemType(underlying, "TimeOnly"))
            {
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.TimeOnlySerial", "global::System.TimeOnly", false);
            }
            if (underlying.TypeKind == TypeKind.Enum)
            {
                string enumType = underlying.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                return ($"global::ExcelReader.Core.Parser.ExcelCellReaders.Enum<{enumType}>", enumType, false);
            }
            if (IsSystemType(underlying, "Guid"))
            {
                return (string.Empty, "global::System.Guid", true);
            }
            if (IsUtf8SpanParsable(underlying))
            {
                string t = underlying.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                return ($"global::ExcelReader.Core.Parser.ExcelCellReaders.Parsable<{t}>", t, false);
            }
            return null;
        }

        private static bool TryGetNullableUnderlying(ITypeSymbol type, out ITypeSymbol underlying)
        {
            if (type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } named)
            {
                underlying = named.TypeArguments[0];
                return true;
            }
            underlying = type;
            return false;
        }

        private static bool IsSystemType(ITypeSymbol type, string name)
        {
            return type is INamedTypeSymbol { ContainingNamespace: { IsGlobalNamespace: false } ns } named
                && string.Equals(ns.ToDisplayString(), "System", StringComparison.Ordinal)
                && string.Equals(named.Name, name, StringComparison.Ordinal);
        }

        private static bool IsUtf8SpanParsable(ITypeSymbol type)
        {
            return type.AllInterfaces.Any(i =>
                string.Equals(i.OriginalDefinition.Name, "IUtf8SpanParsable", StringComparison.Ordinal)
                && string.Equals(i.OriginalDefinition.ContainingNamespace?.ToDisplayString(), "System", StringComparison.Ordinal)
                && i.TypeArguments.Length == 1
                && SymbolEqualityComparer.Default.Equals(i.TypeArguments[0], type));
        }

        // Symbol-level equivalent of "converterType implements openInterface<exactArgument>" — used for
        // both IExcelCellConverter<T> (read) and IExcelCellWriter<T> (write), each checked independently
        // since a converter may implement only one of the two (a read-only converter has no write side).
        private static bool ImplementsGenericInterface(ITypeSymbol converterType, string openInterfaceDisplay, ITypeSymbol exactArgument)
        {
            return converterType.AllInterfaces.Any(i =>
                string.Equals(i.OriginalDefinition.ToDisplayString(), openInterfaceDisplay, StringComparison.Ordinal)
                && i.TypeArguments.Length == 1
                && SymbolEqualityComparer.Default.Equals(i.TypeArguments[0], exactArgument));
        }

        private static string GenerateSource(INamedTypeSymbol symbol, List<PropertyPlan> properties)
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

            foreach (string? decl in properties.Select(static p => p.ConverterFieldDecl).Where(static d => d is not null))
            {
                sb.AppendLine(decl);
            }

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

        private static void AppendRowMap(StringBuilder sb, INamedTypeSymbol symbol, string qualifiedType, List<PropertyPlan> properties)
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
                EmitReadFragment(sb, qualifiedType, p);
            }
            sb.AppendLine("        ;");
            sb.AppendLine("    }");
        }

        private static void EmitReadFragment(StringBuilder sb, string qualifiedType, PropertyPlan p)
        {
            string namesLiteral = string.Join(", ", p.HeaderNames.Select(static n => $"\"{n.Replace("\"", "\\\"")}\""));
            string req = $"isRequired: {Bool(p.IsRequired)}, requireValue: {Bool(p.RequireValue)}";
            switch (p.Read.Kind)
            {
                case ReadKind.None:
                    return;
                case ReadKind.Value:
                    sb.AppendLine($"            .Property([{namesLiteral}], {p.Read.Reader}, static (ref {qualifiedType} m, {p.Read.ValueType} v) => m.{p.PropertyName} = v, {req})");
                    return;
                case ReadKind.Nullable:
                    sb.AppendLine($"            .PropertyNullable([{namesLiteral}], {p.Read.Reader}, static (ref {qualifiedType} m, {p.Read.ValueType}? v) => m.{p.PropertyName} = v, {req})");
                    return;
                case ReadKind.Converted:
                    sb.AppendLine($"            .Converted([{namesLiteral}], {p.Read.Reader}, static (ref {qualifiedType} m, {p.Read.ValueType} v) => m.{p.PropertyName} = v, {req})");
                    return;
                case ReadKind.GuidValue:
                    sb.AppendLine("#if NET9_0_OR_GREATER");
                    sb.AppendLine($"            .Property([{namesLiteral}], global::ExcelReader.Core.Parser.ExcelCellReaders.Parsable<global::System.Guid>, static (ref {qualifiedType} m, global::System.Guid v) => m.{p.PropertyName} = v, {req})");
                    sb.AppendLine("#else");
                    sb.AppendLine($"            .Property([{namesLiteral}], global::ExcelReader.Core.Parser.ExcelCellReaders.Guid, static (ref {qualifiedType} m, global::System.Guid v) => m.{p.PropertyName} = v, {req})");
                    sb.AppendLine("#endif");
                    return;
                case ReadKind.GuidNullable:
                    sb.AppendLine("#if NET9_0_OR_GREATER");
                    sb.AppendLine($"            .PropertyNullable([{namesLiteral}], global::ExcelReader.Core.Parser.ExcelCellReaders.Parsable<global::System.Guid>, static (ref {qualifiedType} m, global::System.Guid? v) => m.{p.PropertyName} = v, {req})");
                    sb.AppendLine("#else");
                    sb.AppendLine($"            .PropertyNullable([{namesLiteral}], global::ExcelReader.Core.Parser.ExcelCellReaders.Guid, static (ref {qualifiedType} m, global::System.Guid? v) => m.{p.PropertyName} = v, {req})");
                    sb.AppendLine("#endif");
                    return;
                default:
                    return;
            }
        }

        private static void AppendRecordMap(StringBuilder sb, string qualifiedType, List<PropertyPlan> properties)
        {
            sb.AppendLine($"    public static void ConfigureExcelRecordMap(global::ExcelReader.Core.Writer.ExcelRecordMapBuilder<{qualifiedType}> builder)");
            sb.AppendLine("    {");
            sb.AppendLine("        builder");
            foreach (string? writeEmit in properties.Select(static p => p.WriteEmit).Where(static w => w is not null))
            {
                sb.AppendLine(writeEmit);
            }
            sb.AppendLine("        ;");
            sb.AppendLine("    }");
        }

        private static string Bool(bool value)
        {
            return value ? "true" : "false";
        }

        // netstandard2.0 has no record types (IsExternalInit lives in netstandard2.1+), so plain
        // constructor-initialized structs do the same job as record struct here.
        private readonly struct ReadPlan
        {
            internal ReadPlan(ReadKind kind, string reader, string valueType)
            {
                Kind = kind;
                Reader = reader;
                ValueType = valueType;
            }

            internal ReadKind Kind { get; }
            internal string Reader { get; }
            internal string ValueType { get; }
        }

        private readonly struct PropertyPlan
        {
            internal PropertyPlan(string propertyName, string[] headerNames, ReadPlan read, bool isRequired, bool requireValue, string? writeEmit, string? converterFieldDecl)
            {
                PropertyName = propertyName;
                HeaderNames = headerNames;
                Read = read;
                IsRequired = isRequired;
                RequireValue = requireValue;
                WriteEmit = writeEmit;
                ConverterFieldDecl = converterFieldDecl;
            }

            internal string PropertyName { get; }
            internal string[] HeaderNames { get; }
            internal ReadPlan Read { get; }
            internal bool IsRequired { get; }
            internal bool RequireValue { get; }
            internal string? WriteEmit { get; }
            internal string? ConverterFieldDecl { get; }
        }
    }
}
