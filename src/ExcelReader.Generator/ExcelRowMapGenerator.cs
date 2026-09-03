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
            "Header name '{0}' is bound by more than one property of '{1}'; only the first one encountered will ever match",
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
            "Type '{0}' is generic; [ExcelSerializable] supports only non-generic types. Map it with ExcelFluentParser<T> or a hand-written IExcelRowMap<T> instead.",
            "ExcelReader.Generator",
            DiagnosticSeverity.Error,
            isEnabledByDefault: true);

        private static readonly DiagnosticDescriptor NoParameterlessConstructorDescriptor = new(
            "EXR008",
            "Type has no public parameterless constructor",
            "Type '{0}' has no public parameterless constructor, so no row instance can be created for it; add one, or map it with ExcelFluentParser<T>",
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
            IncrementalValuesProvider<GeneratedResult> results = context.SyntaxProvider
                .ForAttributeWithMetadataName(
                    SerializableAttribute,
                    // TypeDeclarationSyntax covers class/struct/record/record struct (but not enum,
                    // which is a BaseTypeDeclarationSyntax, not a TypeDeclarationSyntax) — a record was
                    // previously silently ignored here (ClassDeclarationSyntax does not cover it).
                    predicate: static (node, _) => node is TypeDeclarationSyntax and not InterfaceDeclarationSyntax,
                    transform: static (ctx, _) => Analyze((INamedTypeSymbol)ctx.TargetSymbol));

            // The transform above does the actual analysis and returns a plain-data, value-equatable
            // result — RegisterSourceOutput's callback only replays it. ForAttributeWithMetadataName's
            // incremental cache compares that result to the previous run's by Equals: an ISymbol/
            // Compilation carried past the transform stage is reference-equality-only across
            // compilations, so it would never hit the cache and this generator would re-run on every
            // keystroke, not just when a marked type's own declaration actually changes.
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
            // A generic hintName (AddSource's file name) is invalid, and the emitted declaration below
            // has no way to carry the type parameters back — reject up front instead of letting AddSource
            // throw (which the generator host reports as an opaque CS8785 with no diagnostic at all).
            if (IsGenericTypeOrContainer(symbol))
            {
                diagnostics.Add(DiagnosticInfo.Create(GenericTypeNotSupportedDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name));
                return NoSource(diagnostics);
            }
            // A struct always has an implicit parameterless constructor; only a class/record class needs
            // one to exist explicitly for `new T()` (AppendRowMap's factory) to compile.
            if (symbol.TypeKind != TypeKind.Struct && !HasPublicParameterlessConstructor(symbol))
            {
                diagnostics.Add(DiagnosticInfo.Create(NoParameterlessConstructorDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name));
                return NoSource(diagnostics);
            }

            IPropertySymbol[] candidateProperties = [.. CollectMappableProperties(symbol)];

            var boundHeaders = new HashSet<string>(StringComparer.Ordinal);
            var plans = new List<PropertyPlan>(candidateProperties.Length);
            bool hadPropertyError = false;
            foreach (IPropertySymbol property in candidateProperties)
            {
                PropertyPlan plan = BuildPlan(diagnostics, symbol, property, ref hadPropertyError);
                if (plan.Read.Kind != ReadKind.None)
                {
                    foreach (string name in plan.HeaderNames.Where(n => !boundHeaders.Add(n)))
                    {
                        diagnostics.Add(DiagnosticInfo.Create(DuplicateHeaderDescriptor, symbol.Locations.FirstOrDefault(), name, symbol.Name));
                    }
                }
                if (plan.Read.Kind != ReadKind.None || plan.WriteEmit is not null)
                {
                    plans.Add(plan);
                }
            }

            // Skip EXR005 when a more specific per-property error (EXR003/EXR004) already explains why
            // nothing mapped — piling a generic "nothing mappable" warning on top of the real cause is
            // noise, not information.
            if (plans.Count == 0 && !hadPropertyError)
            {
                diagnostics.Add(DiagnosticInfo.Create(NoMappablePropertyDescriptor, symbol.Locations.FirstOrDefault(), symbol.Name));
            }

            string source = GenerateSource(symbol, plans);
            string hintName = symbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat)
                .Replace("global::", "")
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

        // Mirrors TypeMapper<T>.Build()'s GetProperties(BindingFlags.Public | BindingFlags.Instance),
        // which (unlike symbol.GetMembers()) walks inherited properties too — a property declared only
        // on a base type was previously missing from the generated map (divergence from reflection).
        // Declared-on-`symbol` properties come first, then each base type's in turn; a property that a
        // derived type re-declares (`new`) shadows the base one, first occurrence by name wins.
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

            // An init-only setter (`{ get; init; }`, and every positional-record property) can be
            // assigned through reflection's CreateDelegate, but the emitted `m.Prop = v` lambda cannot
            // (CS8852) — such a property is write-only from the generator's perspective, so it's read
            // through the record's write side only, not the row map.
            bool canSet = property.SetMethod is { DeclaredAccessibility: Accessibility.Public, IsInitOnly: false };
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
                    diagnostics.Add(DiagnosticInfo.Create(
                        ConverterTypeMismatchDescriptor, property.Locations.FirstOrDefault(), converterQualified, qualifiedProperty, owner.Name, property.Name));
                    hadPropertyError = true;
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
                    diagnostics.Add(DiagnosticInfo.Create(
                        RequiredWithNoParserDescriptor, property.Locations.FirstOrDefault(), owner.Name, property.Name, qualifiedProperty));
                    hadPropertyError = true;
                }

                WriteKind writeKind = GetWriteKind(underlying);
                if (canGet && writeKind != WriteKind.None)
                {
                    // A Nullable<T> (isNullable) or a reference type (!underlying.IsValueType) can be
                    // null at runtime — matches RecordColumns<T>.ToStringExpression's reflection-path
                    // rule exactly (Plan<TRow>.Build in WorkbookRecordWriter.cs): a genuinely non-null
                    // value type calls ToString() directly, everything else null-conditionally.
                    bool needsNullConditional = isNullable || !underlying.IsValueType;
                    string valueExpr = WriteValueExpression(writeKind, needsNullConditional, property.Name);
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

        private static string WriteValueExpression(WriteKind kind, bool needsNullConditional, string propertyName)
        {
            if (kind == WriteKind.Direct)
            {
                return $"m.{propertyName}";
            }
            return needsNullConditional ? $"m.{propertyName}?.ToString()" : $"m.{propertyName}.ToString()";
        }

        // Direct for the fixed set RowWriteMethods<TRow> has a dedicated Write overload for; every
        // other type — enum, Guid, char, Half, Int128, UInt128, TimeSpan, DateTimeOffset, a plain
        // reference type with no built-in reader, anything — falls back to ToString(), exactly like
        // RecordColumns<T>.Plan<TRow>.Build's reflection path (RowWriteMethods<TRow>.Select's
        // asString=true default). No type is ever omitted from the written columns.
        private static WriteKind GetWriteKind(ITypeSymbol underlying)
        {
            if (underlying.SpecialType is SpecialType.System_String or SpecialType.System_Boolean
                || IsSystemType(underlying, "DateTime") || IsSystemType(underlying, "DateOnly") || IsSystemType(underlying, "TimeOnly")
                || NumericWriteSpecialTypes.Contains(underlying.SpecialType))
            {
                return WriteKind.Direct;
            }
            return WriteKind.ToStringFallback;
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
            // *Auto, not *Serial: ExcelMappedParser<T> builds one map and reuses it for every reader
            // (unlike the reflection path's dedicated csvTextDates map for CSV), so a date property has
            // to read a serial number from XLSX/XLSB/XLS and ISO text from CSV through the same reader.
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

        // Matches the original declaration's keyword so the emitted partial declares the same kind —
        // "class" for a record class would compile but subtly change the type's semantics (no more
        // synthesized Equals/ToString from a *second* declaration, though the original still has them),
        // whereas "struct" for a record struct's partial would flat out fail to bind IsRecord's members.
        private static string DeclarationKeyword(INamedTypeSymbol symbol)
        {
            if (symbol.TypeKind == TypeKind.Struct)
            {
                return symbol.IsRecord ? "record struct" : "struct";
            }
            return symbol.IsRecord ? "record" : "class";
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
                sb.AppendLine($"partial {DeclarationKeyword(outer)} {outer.Name}");
                sb.AppendLine("{");
            }

            sb.AppendLine($"partial {DeclarationKeyword(symbol)} {symbol.Name} : global::ExcelReader.Core.Parser.IExcelRowMap<{qualifiedType}>, global::ExcelReader.Core.Writer.IExcelRecordMap<{qualifiedType}>");
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

        // Emits a single fused .PropertyRaw(...) call per bound property instead of the two-delegate
        // .Property/.PropertyNullable/.Converted composition ExcelRowMapBuilder<T> builds internally:
        // one indirect call per bound cell per row (the ColumnParser<T> itself) instead of two (the
        // ColumnParser<T> wrapper calling through to a separate read delegate, then a separate setter
        // delegate). Assigning a non-nullable `v` straight into a `TValue?` property (the Nullable/
        // GuidNullable cases) already implicitly wraps it, so those collapse into the same emission as
        // their non-nullable counterparts — PropertyRaw needs no separate nullable overload the way
        // Property/PropertyNullable do.
        private static void EmitReadFragment(StringBuilder sb, string qualifiedType, PropertyPlan p)
        {
            string namesLiteral = string.Join(", ", p.HeaderNames.Select(static n => $"\"{n.Replace("\"", "\\\"")}\""));
            string req = $"isRequired: {Bool(p.IsRequired)}, requireValue: {Bool(p.RequireValue)}";
            switch (p.Read.Kind)
            {
                case ReadKind.None:
                    return;
                case ReadKind.Value:
                case ReadKind.Nullable:
                    EmitPropertyRaw(sb, qualifiedType, namesLiteral, p.PropertyName, req,
                        $"{p.Read.Reader}(in c, d, pr, out {p.Read.ValueType} v)");
                    return;
                case ReadKind.Converted:
                    EmitPropertyRaw(sb, qualifiedType, namesLiteral, p.PropertyName, req,
                        $"{p.Read.Reader}.TryConvert(in c, d, pr, out {p.Read.ValueType} v)");
                    return;
                case ReadKind.GuidValue:
                case ReadKind.GuidNullable:
                    sb.AppendLine("#if NET9_0_OR_GREATER");
                    EmitPropertyRaw(sb, qualifiedType, namesLiteral, p.PropertyName, req,
                        "global::ExcelReader.Core.Parser.ExcelCellReaders.Parsable<global::System.Guid>(in c, d, pr, out global::System.Guid v)");
                    sb.AppendLine("#else");
                    EmitPropertyRaw(sb, qualifiedType, namesLiteral, p.PropertyName, req,
                        "global::ExcelReader.Core.Parser.ExcelCellReaders.Guid(in c, d, pr, out global::System.Guid v)");
                    sb.AppendLine("#endif");
                    return;
                default:
                    return;
            }
        }

        // `tryReadExpr` is a full call expression evaluating a `bool`, binding `out {ValueType} v` — e.g.
        // "global::...ExcelCellReaders.Bool(in c, d, pr, out bool v)" or "s_converter_Foo.TryConvert(in
        // c, d, pr, out int v)". `c`/`d`/`pr` are this lambda's own parameter names (Cell/isDate1904/
        // provider), matching ExcelRowParser<T>'s shape exactly so no wrapper indirection is introduced.
        private static void EmitPropertyRaw(StringBuilder sb, string qualifiedType, string namesLiteral, string propertyName, string req, string tryReadExpr)
        {
            sb.AppendLine($"            .PropertyRaw([{namesLiteral}], static (ref {qualifiedType} m, in global::ExcelReader.Core.ValueObjects.Cell c, bool d, global::System.IFormatProvider pr) =>");
            sb.AppendLine("            {");
            sb.AppendLine($"                if (!{tryReadExpr}) {{ return false; }}");
            sb.AppendLine($"                m.{propertyName} = v;");
            sb.AppendLine("                return true;");
            sb.AppendLine($"            }}, {req})");
        }

        private static void AppendRecordMap(StringBuilder sb, string qualifiedType, List<PropertyPlan> properties)
        {
            // Generic in TRow, matching IExcelRecordMap<T>: the body is the same for every row writer,
            // but each instantiation binds its column actions to that concrete writer instead of to
            // IRowWriter, so a cell write does not dispatch through the interface.
            sb.AppendLine($"    public static void ConfigureExcelRecordMap<TRow>(global::ExcelReader.Core.Writer.ExcelRecordMapBuilder<{qualifiedType}, TRow> builder)");
            sb.AppendLine("        where TRow : global::ExcelReader.Core.Writer.IRowWriter");
            sb.AppendLine("    {");
            string[] writeEmits = [.. properties.Select(static p => p.WriteEmit).Where(static w => w is not null)!];
            if (writeEmits.Length == 0)
            {
                // No `.Column(...)` calls to chain — `builder;` alone is not a valid statement
                // (CS0201), and an unused parameter warning would fire without some use of it.
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

        // Everything Analyze() produces for one [ExcelSerializable] type, carrying no ISymbol/
        // Compilation/live Location past the transform stage — plain strings and EquatableArray, so two
        // runs whose underlying declaration didn't change compare Equals and RegisterSourceOutput's
        // callback is skipped for that type, instead of re-running on every keystroke regardless of
        // whether anything relevant changed. A plain struct, not a record: netstandard2.0 (this
        // project's required TFM, see the .csproj comment) has no IsExternalInit, so record/record
        // struct's compiler-synthesized init-setters don't compile here — same reason ReadPlan/
        // PropertyPlan above are plain structs.
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

        // Location.Create(filePath, textSpan, lineSpan) — the "external file" flavor — captures the
        // same information a real syntax Location does, but as plain value data with no reference to
        // the SyntaxTree/Compilation that produced it, which is what makes it safe to compare across
        // incremental runs. The real, tree-bound Location that comes out of the symbol API is exactly
        // the kind of thing that would otherwise pin the whole Analyze() result to reference equality.
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

        // ImmutableArray<T>'s own Equals is reference-equality on the backing array, not structural —
        // exactly wrong for an incremental-generator cache key, where two runs' arrays are never the
        // same instance even when their content is identical. This wrapper is the standard fix.
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

        // System.HashCode isn't available on netstandard2.0 (this project's required TFM) — the
        // classic combine formula does the same job without it.
        private static int CombineHash(int h1, int h2)
        {
            unchecked
            {
                return (h1 * 397) ^ h2;
            }
        }
    }
}
