using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

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
            "Header name '{0}' is bound by more than one property of '{1}' under the default case-insensitive, trimmed header matching; parsing with that matching throws InvalidOperationException",
            "ExcelReader.Generator",
            DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        private static readonly DiagnosticDescriptor NoMappablePropertyDescriptor = new(
            "EXR005",
            "Type marked [ExcelSerializable] has no mappable property",
            "Type '{0}' is marked [ExcelSerializable] but has no property that can be read or written; the generated map is empty",
            "ExcelReader.Generator",
            DiagnosticSeverity.Warning,
            isEnabledByDefault: true);

        private static readonly DiagnosticDescriptor GenericTypeNotSupportedDescriptor = new(
            "EXR007",
            "[ExcelSerializable] does not support generic types",
            "Type '{0}' is generic; [ExcelSerializable] supports only non-generic types. Map it with ExcelParser.Build or a hand-written IExcelRowMap<T> instead.",
            "ExcelReader.Generator",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        private static readonly DiagnosticDescriptor NoParameterlessConstructorDescriptor = new(
            "EXR008",
            "Type has no public parameterless constructor",
            "Type '{0}' has no public parameterless constructor, so no row instance can be created for it; add one, or map it with ExcelParser.Build",
            "ExcelReader.Generator",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        private static readonly DiagnosticDescriptor RequiredWithNoSetterDescriptor = new(
            "EXR009",
            "[ExcelRequired] property has no public setter",
            "Property '{0}.{1}' is marked [ExcelRequired] but has no public set or init accessor, so it can never be read; add one, or remove [ExcelRequired]",
            "ExcelReader.Generator",
            DiagnosticSeverity.Error,
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
            InvariantText,
            Utf8Text,
        }

        private static readonly HashSet<SpecialType> NumericWriteSpecialTypes =
        [
            SpecialType.System_Byte, SpecialType.System_SByte, SpecialType.System_Int16, SpecialType.System_UInt16,
            SpecialType.System_Int32, SpecialType.System_UInt32, SpecialType.System_Int64, SpecialType.System_UInt64,
            SpecialType.System_Single, SpecialType.System_Double, SpecialType.System_Decimal,
        ];

        /// <inheritdoc/>
        public void Initialize(IncrementalGeneratorInitializationContext context)
        {
            IncrementalValuesProvider<GeneratedResult> results = context.SyntaxProvider
                .ForAttributeWithMetadataName(
                    SerializableAttribute,
                    predicate: static (node, _) => node is TypeDeclarationSyntax and not InterfaceDeclarationSyntax,
                    transform: static (ctx, _) => Analyze((INamedTypeSymbol)ctx.TargetSymbol));

            context.RegisterSourceOutput(results, static (spc, result) =>
            {
                foreach (DiagnosticInfo diagnostic in result.Diagnostics.Items)
                {
                    spc.ReportDiagnostic(diagnostic.ToDiagnostic());
                }
                if (result.Source is not null)
                {
                    spc.AddSource($"{result.HintName}.ExcelRowMap.g.cs", result.Source);
                }
            });
        }

        private static GeneratedResult Analyze(INamedTypeSymbol symbol)
        {
            var diagnostics = new List<DiagnosticInfo>();
            if (!IsPartial(symbol))
            {
                diagnostics.Add(DiagnosticInfo.Create(NotPartialDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name));
                return NoSource(diagnostics);
            }
            for (INamedTypeSymbol? outer = symbol.ContainingType; outer is not null; outer = outer.ContainingType)
            {
                if (!IsPartial(outer))
                {
                    diagnostics.Add(DiagnosticInfo.Create(ContainingTypeNotPartialDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name, outer.Name));
                    return NoSource(diagnostics);
                }
            }
            if (IsGenericTypeOrContainer(symbol))
            {
                diagnostics.Add(DiagnosticInfo.Create(GenericTypeNotSupportedDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name));
                return NoSource(diagnostics);
            }
            if (symbol.TypeKind != TypeKind.Struct && !HasPublicParameterlessConstructor(symbol))
            {
                diagnostics.Add(DiagnosticInfo.Create(NoParameterlessConstructorDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name));
                return NoSource(diagnostics);
            }

            IPropertySymbol[] candidateProperties = [.. CollectMappableProperties(symbol)];

            var headerOwners = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var plans = new List<PropertyPlan>(candidateProperties.Length);
            bool hadPropertyError = false;
            foreach (IPropertySymbol property in candidateProperties)
            {
                PropertyPlan plan = BuildPlan(diagnostics, symbol, property, ref hadPropertyError);
                if (plan.Read.Kind != ReadKind.None)
                {
                    foreach (string name in plan.HeaderNames)
                    {
                        string key = name.Trim();
                        if (!headerOwners.TryGetValue(key, out string? owner))
                        {
                            headerOwners.Add(key, property.Name);
                        }
                        else if (!string.Equals(owner, property.Name, StringComparison.Ordinal))
                        {
                            diagnostics.Add(DiagnosticInfo.Create(DuplicateHeaderDescriptor, symbol.Locations.FirstOrDefault(), name, symbol.Name));
                        }
                    }
                }
                if (plan.Read.Kind != ReadKind.None || plan.WriteEmit is not null)
                {
                    plans.Add(plan);
                }
            }

            if (plans.Count == 0 && !hadPropertyError)
            {
                diagnostics.Add(DiagnosticInfo.Create(NoMappablePropertyDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name));
            }

            string source = GenerateSource(symbol, plans);
            string hintName = symbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat)
                .Replace("global::", "")
                .Replace("@", "")
                .Replace('.', '_');
            return new GeneratedResult(hintName, source, new EquatableArray<DiagnosticInfo>([.. diagnostics]));
        }

        private static GeneratedResult NoSource(List<DiagnosticInfo> diagnostics)
        {
            return new GeneratedResult(null, null, new EquatableArray<DiagnosticInfo>([.. diagnostics]));
        }

        private static bool IsGenericTypeOrContainer(INamedTypeSymbol symbol)
        {
            if (symbol.TypeParameters.Length > 0)
            {
                return true;
            }
            for (INamedTypeSymbol? outer = symbol.ContainingType; outer is not null; outer = outer.ContainingType)
            {
                if (outer.TypeParameters.Length > 0)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool HasPublicParameterlessConstructor(INamedTypeSymbol symbol)
        {
            return symbol.Constructors.Any(static c => c.Parameters.Length == 0 && c.DeclaredAccessibility == Accessibility.Public);
        }

        private static IEnumerable<IPropertySymbol> CollectMappableProperties(INamedTypeSymbol symbol)
        {
            var seenNames = new HashSet<string>(StringComparer.Ordinal);
            for (INamedTypeSymbol? current = symbol; current is not null && current.SpecialType != SpecialType.System_Object; current = current.BaseType)
            {
                foreach (IPropertySymbol property in current.GetMembers().OfType<IPropertySymbol>()
                    .Where(p => !p.IsStatic && p.Parameters.Length == 0 && p.DeclaredAccessibility == Accessibility.Public
                        && !HasAttribute(p, IgnoreAttribute) && seenNames.Add(p.Name)))
                {
                    yield return property;
                }
            }
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

        private static PropertyPlan BuildPlan(List<DiagnosticInfo> diagnostics, INamedTypeSymbol owner, IPropertySymbol property, ref bool hadPropertyError)
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

            string member = Identifier(property.Name);
            bool canSet = property.SetMethod is { DeclaredAccessibility: Accessibility.Public };
            bool canGet = property.GetMethod is { DeclaredAccessibility: Accessibility.Public };
            string qualifiedProperty = property.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
            if (isRequired && !canSet)
            {
                diagnostics.Add(DiagnosticInfo.Create(
                    RequiredWithNoSetterDescriptor, property.Locations.FirstOrDefault(), owner.Name, property.Name));
                hadPropertyError = true;
            }

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
                    diagnostics.Add(DiagnosticInfo.Create(
                        ConverterTypeMismatchDescriptor, property.Locations.FirstOrDefault(), converterQualified, qualifiedProperty, owner.Name, property.Name));
                    hadPropertyError = true;
                }
                if (implementsWriter && canGet)
                {
                    writeEmit = $"            .Column({Literal(names[0])}, static (row, m) => {fieldName}.Write(row, m.{member}))";
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
                else if (isRequired && canSet)
                {
                    diagnostics.Add(DiagnosticInfo.Create(
                        RequiredWithNoParserDescriptor, property.Locations.FirstOrDefault(), owner.Name, property.Name, qualifiedProperty));
                    hadPropertyError = true;
                }

                WriteKind writeKind = GetWriteKind(underlying);
                if (canGet && writeKind != WriteKind.None)
                {
                    bool needsNullConditional = isNullable || !underlying.IsValueType;
                    string valueExpr = WriteValueExpression(writeKind, needsNullConditional, member, underlying);
                    writeEmit = $"            .Column({Literal(names[0])}, static (row, m) => row.Write({valueExpr}))";
                }
            }

            (string assign, string? initAccessorDecl) = read.Kind == ReadKind.None
                ? ($"m.{member} = v", null)
                : PlanAssignment(property, member, qualifiedProperty);
            return new PropertyPlan(names, read, isRequired, requireValue, writeEmit, converterFieldDecl, assign, initAccessorDecl);
        }

        private static (string Assign, string? InitAccessorDecl) PlanAssignment(IPropertySymbol property, string member, string qualifiedProperty)
        {
            if (property.SetMethod is not { IsInitOnly: true } init)
            {
                return ($"m.{member} = v", null);
            }
            // Init accessors can't be called outside an object initializer; UnsafeAccessor binds the same setter reflection's GetSetMethod() does.
            string accessor = $"__ExcelInit_{property.Name}";
            string byRef = init.ContainingType.IsValueType ? "ref " : "";
            string target = init.ContainingType.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
            string decl = $"    [global::System.Runtime.CompilerServices.UnsafeAccessor(global::System.Runtime.CompilerServices.UnsafeAccessorKind.Method, Name = {Literal(init.MetadataName)})] private static extern void {accessor}({byRef}{target} target, {qualifiedProperty} value);";
            return ($"{accessor}({byRef}m, v)", decl);
        }

        private static string Literal(string value)
        {
            return SymbolDisplay.FormatLiteral(value, quote: true);
        }

        private static string Identifier(string name)
        {
            return SyntaxFacts.GetKeywordKind(name) == SyntaxKind.None ? name : "@" + name;
        }

        private static ReadKind SelectReadKind(bool isGuid, bool isNullable)
        {
            if (isGuid)
            {
                return isNullable ? ReadKind.GuidNullable : ReadKind.GuidValue;
            }
            return isNullable ? ReadKind.Nullable : ReadKind.Value;
        }

        private static string WriteValueExpression(WriteKind kind, bool needsNullConditional, string propertyName, ITypeSymbol underlying)
        {
            if (kind == WriteKind.Utf8Text)
            {
                return $"global::System.Text.Encoding.UTF8.GetString(m.{propertyName})";
            }
            if (kind == WriteKind.Direct)
            {
                return $"m.{propertyName}";
            }
            if (kind == WriteKind.InvariantText)
            {
                string args = (IsSystemType(underlying, "DateTimeOffset") ? "\"O\"" : "null") + ", global::System.Globalization.CultureInfo.InvariantCulture";
                if (HasPublicFormattableToString(underlying))
                {
                    return needsNullConditional ? $"m.{propertyName}?.ToString({args})" : $"m.{propertyName}.ToString({args})";
                }
                if (!underlying.IsValueType)
                {
                    return $"((global::System.IFormattable?)m.{propertyName})?.ToString({args})";
                }
                return needsNullConditional
                    ? $"(m.{propertyName} is {{ }} v ? ((global::System.IFormattable)v).ToString({args}) : null)"
                    : $"((global::System.IFormattable)m.{propertyName}).ToString({args})";
            }
            return needsNullConditional ? $"m.{propertyName}?.ToString()" : $"m.{propertyName}.ToString()";
        }

        private static bool ImplementsFormattable(ITypeSymbol type)
        {
            return type.AllInterfaces.Any(static i => string.Equals(i.ToDisplayString(), "System.IFormattable", StringComparison.Ordinal));
        }

        private static bool HasPublicFormattableToString(ITypeSymbol type)
        {
            return type.GetMembers("ToString").OfType<IMethodSymbol>().Any(static m =>
                m is { DeclaredAccessibility: Accessibility.Public, IsStatic: false, Parameters.Length: 2 }
                && m.Parameters[0].Type.SpecialType == SpecialType.System_String
                && string.Equals(m.Parameters[1].Type.ToDisplayString(), "System.IFormatProvider", StringComparison.Ordinal));
        }

        private static WriteKind GetWriteKind(ITypeSymbol underlying)
        {
            if (IsUtf8Span(underlying))
            {
                return WriteKind.Utf8Text;
            }
            if (underlying.SpecialType is SpecialType.System_String or SpecialType.System_Boolean
                || IsSystemType(underlying, "DateTime") || IsSystemType(underlying, "DateOnly") || IsSystemType(underlying, "TimeOnly")
                || IsSystemType(underlying, "Half") || NumericWriteSpecialTypes.Contains(underlying.SpecialType))
            {
                return WriteKind.Direct;
            }
            if (underlying.TypeKind != TypeKind.Enum && ImplementsFormattable(underlying))
            {
                return WriteKind.InvariantText;
            }
            return WriteKind.ToStringFallback;
        }

        private static (string Reader, string ValueType, bool IsGuid)? TryGetBuiltInReader(ITypeSymbol underlying)
        {
            if (IsUtf8Span(underlying))
            {
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.Utf8", "global::System.ReadOnlySpan<byte>", false);
            }
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
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.DateTimeAuto", "global::System.DateTime", false);
            }
            if (IsSystemType(underlying, "DateOnly"))
            {
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.DateOnlyAuto", "global::System.DateOnly", false);
            }
            if (IsSystemType(underlying, "TimeOnly"))
            {
                return ("global::ExcelReader.Core.Parser.ExcelCellReaders.TimeOnlyAuto", "global::System.TimeOnly", false);
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
            if (IsSelfParsable(underlying, "IUtf8SpanParsable"))
            {
                string t = underlying.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                return ($"global::ExcelReader.Core.Parser.ExcelCellReaders.Parsable<{t}>", t, false);
            }
            if (IsSelfParsable(underlying, "ISpanParsable"))
            {
                string t = underlying.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                return ($"global::ExcelReader.Core.Parser.ExcelCellReaders.SpanParsable<{t}>", t, false);
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

        private static bool IsUtf8Span(ITypeSymbol type)
        {
            return type is INamedTypeSymbol { Name: "ReadOnlySpan", TypeArguments.Length: 1 } named
                && named.TypeArguments[0].SpecialType == SpecialType.System_Byte
                && string.Equals(named.ContainingNamespace?.ToDisplayString(), "System", StringComparison.Ordinal);
        }

        private static bool IsSystemType(ITypeSymbol type, string name)
        {
            return type is INamedTypeSymbol { ContainingNamespace: { IsGlobalNamespace: false } ns } named
                && string.Equals(ns.ToDisplayString(), "System", StringComparison.Ordinal)
                && string.Equals(named.Name, name, StringComparison.Ordinal);
        }

        private static bool IsSelfParsable(ITypeSymbol type, string interfaceName)
        {
            return type.AllInterfaces.Any(i =>
                string.Equals(i.OriginalDefinition.Name, interfaceName, StringComparison.Ordinal)
                && string.Equals(i.OriginalDefinition.ContainingNamespace?.ToDisplayString(), "System", StringComparison.Ordinal)
                && i.TypeArguments.Length == 1
                && SymbolEqualityComparer.Default.Equals(i.TypeArguments[0], type));
        }

        private static bool ImplementsGenericInterface(ITypeSymbol converterType, string openInterfaceDisplay, ITypeSymbol exactArgument)
        {
            return converterType.AllInterfaces.Any(i =>
                string.Equals(i.OriginalDefinition.ToDisplayString(), openInterfaceDisplay, StringComparison.Ordinal)
                && i.TypeArguments.Length == 1
                && SymbolEqualityComparer.Default.Equals(i.TypeArguments[0], exactArgument));
        }

        private static string PartialDeclaration(INamedTypeSymbol symbol)
        {
            if (symbol.TypeKind == TypeKind.Struct)
            {
                if (symbol.IsRefLikeType)
                {
                    return "ref partial struct";
                }
                return symbol.IsRecord ? "partial record struct" : "partial struct";
            }
            return symbol.IsRecord ? "partial record" : "partial class";
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
                sb.AppendLine($"{PartialDeclaration(outer)} {Identifier(outer.Name)}");
                sb.AppendLine("{");
            }

            sb.AppendLine($"{PartialDeclaration(symbol)} {Identifier(symbol.Name)} : global::ExcelReader.Core.Parser.IExcelRowMap<{qualifiedType}>, global::ExcelReader.Core.Writer.IExcelRecordMap<{qualifiedType}>");
            sb.AppendLine("{");

            foreach (string? decl in properties.SelectMany(static p => new[] { p.ConverterFieldDecl, p.InitAccessorDecl }).Where(static d => d is not null))
            {
                sb.AppendLine(decl);
            }

            AppendRowMap(sb, symbol, qualifiedType, properties);
            AppendRecordMap(sb, qualifiedType, properties);

            sb.AppendLine("}");
            for (int i = 0; i < containers.Count; i++)
            {
                sb.AppendLine("}");
            }
            if (ns is not null)
            {
                sb.AppendLine("}");
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
            string namesLiteral = string.Join(", ", p.HeaderNames.Select(Literal));
            string req = $"isRequired: {Bool(p.IsRequired)}, requireValue: {Bool(p.RequireValue)}";
            switch (p.Read.Kind)
            {
                case ReadKind.None:
                    return;
                case ReadKind.Value:
                case ReadKind.Nullable:
                    EmitPropertyRaw(sb, qualifiedType, namesLiteral, p.Assign, req,
                        $"{p.Read.Reader}(in c, d, pr, out {p.Read.ValueType} v)");
                    return;
                case ReadKind.Converted:
                    EmitPropertyRaw(sb, qualifiedType, namesLiteral, p.Assign, req,
                        $"{p.Read.Reader}.TryConvert(in c, d, pr, out {p.Read.ValueType} v)");
                    return;
                case ReadKind.GuidValue:
                case ReadKind.GuidNullable:
                    EmitPropertyRaw(sb, qualifiedType, namesLiteral, p.Assign, req,
                        "global::ExcelReader.Core.Parser.ExcelCellReaders.Parsable<global::System.Guid>(in c, d, pr, out global::System.Guid v)");
                    return;
                default:
                    return;
            }
        }

        private static void EmitPropertyRaw(StringBuilder sb, string qualifiedType, string namesLiteral, string assign, string req, string tryReadExpr)
        {
            sb.AppendLine($"            .PropertyRaw([{namesLiteral}], static (ref {qualifiedType} m, in global::ExcelReader.Core.Reader.Cell c, bool d, global::System.IFormatProvider pr) =>");
            sb.AppendLine("            {");
            sb.AppendLine($"                if (!{tryReadExpr}) {{ return false; }}");
            sb.AppendLine($"                {assign};");
            sb.AppendLine("                return true;");
            sb.AppendLine($"            }}, {req})");
        }

        private static void AppendRecordMap(StringBuilder sb, string qualifiedType, List<PropertyPlan> properties)
        {
            sb.AppendLine($"    public static void ConfigureExcelRecordMap<TRow>(global::ExcelReader.Core.Writer.ExcelRecordMapBuilder<{qualifiedType}, TRow> builder)");
            sb.AppendLine("        where TRow : global::ExcelReader.Core.Writer.IRowWriter");
            sb.AppendLine("    {");
            string[] writeEmits = [.. properties.Select(static p => p.WriteEmit).Where(static w => w is not null)!];
            if (writeEmits.Length == 0)
            {
                sb.AppendLine("        _ = builder;");
                sb.AppendLine("    }");
                return;
            }
            sb.AppendLine("        builder");
            foreach (string writeEmit in writeEmits)
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
            internal PropertyPlan(string[] headerNames, ReadPlan read, bool isRequired, bool requireValue, string? writeEmit, string? converterFieldDecl, string assign, string? initAccessorDecl)
            {
                HeaderNames = headerNames;
                Read = read;
                IsRequired = isRequired;
                RequireValue = requireValue;
                WriteEmit = writeEmit;
                ConverterFieldDecl = converterFieldDecl;
                Assign = assign;
                InitAccessorDecl = initAccessorDecl;
            }

            internal string[] HeaderNames { get; }
            internal ReadPlan Read { get; }
            internal bool IsRequired { get; }
            internal bool RequireValue { get; }
            internal string? WriteEmit { get; }
            internal string? ConverterFieldDecl { get; }
            internal string Assign { get; }
            internal string? InitAccessorDecl { get; }
        }

        private readonly struct GeneratedResult : IEquatable<GeneratedResult>
        {
            internal GeneratedResult(string? hintName, string? source, EquatableArray<DiagnosticInfo> diagnostics)
            {
                HintName = hintName;
                Source = source;
                Diagnostics = diagnostics;
            }

            internal string? HintName { get; }
            internal string? Source { get; }
            internal EquatableArray<DiagnosticInfo> Diagnostics { get; }

            public bool Equals(GeneratedResult other)
            {
                return string.Equals(HintName, other.HintName, StringComparison.Ordinal)
                    && string.Equals(Source, other.Source, StringComparison.Ordinal)
                    && Diagnostics.Equals(other.Diagnostics);
            }

            public override bool Equals(object? obj)
            {
                return obj is GeneratedResult other && Equals(other);
            }

            public override int GetHashCode()
            {
                int hintHash = HintName is null ? 0 : StringComparer.Ordinal.GetHashCode(HintName);
                int sourceHash = Source is null ? 0 : StringComparer.Ordinal.GetHashCode(Source);
                return CombineHash(CombineHash(hintHash, sourceHash), Diagnostics.GetHashCode());
            }
        }

        private readonly struct LocationInfo : IEquatable<LocationInfo>
        {
            private LocationInfo(string filePath, TextSpan span, LinePositionSpan lineSpan)
            {
                FilePath = filePath;
                Span = span;
                LineSpan = lineSpan;
            }

            private string FilePath { get; }
            private TextSpan Span { get; }
            private LinePositionSpan LineSpan { get; }

            internal Location ToLocation()
            {
                return Location.Create(FilePath, Span, LineSpan);
            }

            internal static LocationInfo? From(Location? location)
            {
                if (location?.SourceTree is null)
                {
                    return null;
                }
                return new LocationInfo(location.SourceTree.FilePath, location.SourceSpan, location.GetLineSpan().Span);
            }

            public bool Equals(LocationInfo other)
            {
                return string.Equals(FilePath, other.FilePath, StringComparison.Ordinal)
                    && Span.Equals(other.Span) && LineSpan.Equals(other.LineSpan);
            }

            public override bool Equals(object? obj)
            {
                return obj is LocationInfo other && Equals(other);
            }

            public override int GetHashCode()
            {
                return CombineHash(CombineHash(StringComparer.Ordinal.GetHashCode(FilePath), Span.GetHashCode()), LineSpan.GetHashCode());
            }
        }

        private readonly struct DiagnosticInfo : IEquatable<DiagnosticInfo>
        {
            private DiagnosticInfo(DiagnosticDescriptor descriptor, LocationInfo? location, EquatableArray<string> messageArgs)
            {
                Descriptor = descriptor;
                Location = location;
                MessageArgs = messageArgs;
            }

            private DiagnosticDescriptor Descriptor { get; }
            private LocationInfo? Location { get; }
            private EquatableArray<string> MessageArgs { get; }

            internal static DiagnosticInfo Create(DiagnosticDescriptor descriptor, Location? location, params string[] messageArgs)
            {
                return new DiagnosticInfo(descriptor, LocationInfo.From(location), new EquatableArray<string>([.. messageArgs]));
            }

            internal Diagnostic ToDiagnostic()
            {
                return Diagnostic.Create(Descriptor, Location?.ToLocation(), [.. MessageArgs.Items]);
            }

            public bool Equals(DiagnosticInfo other)
            {
                return Descriptor.Equals(other.Descriptor) && Nullable.Equals(Location, other.Location) && MessageArgs.Equals(other.MessageArgs);
            }

            public override bool Equals(object? obj)
            {
                return obj is DiagnosticInfo other && Equals(other);
            }

            public override int GetHashCode()
            {
                return CombineHash(Descriptor.GetHashCode(), MessageArgs.GetHashCode());
            }
        }

        private readonly struct EquatableArray<T> : IEquatable<EquatableArray<T>>
            where T : IEquatable<T>
        {
            internal EquatableArray(ImmutableArray<T> items)
            {
                Items = items;
            }

            internal ImmutableArray<T> Items { get; }

            public bool Equals(EquatableArray<T> other)
            {
                if (Items.IsDefault || other.Items.IsDefault)
                {
                    return Items.IsDefault == other.Items.IsDefault;
                }
                if (Items.Length != other.Items.Length)
                {
                    return false;
                }
                for (int i = 0; i < Items.Length; i++)
                {
                    if (!Items[i].Equals(other.Items[i]))
                    {
                        return false;
                    }
                }
                return true;
            }

            public override bool Equals(object? obj)
            {
                return obj is EquatableArray<T> other && Equals(other);
            }

            public override int GetHashCode()
            {
                if (Items.IsDefault)
                {
                    return 0;
                }
                int hash = 17;
                foreach (T item in Items)
                {
                    hash = CombineHash(hash, item.GetHashCode());
                }
                return hash;
            }
        }

        private static int CombineHash(int h1, int h2)
        {
            unchecked
            {
                return (h1 * 397) ^ h2;
            }
        }
    }
}
