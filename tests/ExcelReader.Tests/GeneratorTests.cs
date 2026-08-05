using System.Collections.Immutable;
using System.Linq;
using System.Reflection;
using ExcelReader.Generator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;

namespace ExcelReader.Tests
{
    // Exercises the generator's own plumbing — triggers on [ExcelSerializable], requires 'partial' up
    // the containing-type chain (EXR001/EXR002), and emits code that type-checks against the real public
    // ExcelRowMapBuilder<T>/ExcelRecordMapBuilder<T>/ExcelCellReaders surface. Runs ExcelRowMapGenerator
    // directly via CSharpGeneratorDriver over an in-memory Compilation, rather than through the SDK's
    // live build-time analyzer pipeline — a deliberately broken model (missing 'partial', for the
    // EXR001/EXR002 cases) would otherwise fail this whole project's own build.
    public class GeneratorTests
    {
        // #if NET9_0_OR_GREATER/NET8_0 in generator output (the Guid split) only resolves correctly if
        // the synthetic Compilation defines the same preprocessor symbol the *real* SDK build would —
        // mirror whichever TFM this test assembly itself is running under, rather than leaving it
        // undefined (which would always take the #else branch, regardless of the actual TFM).
#if NET9_0_OR_GREATER
        private static readonly string[] _preprocessorSymbols = ["NET9_0_OR_GREATER"];
#else
        private static readonly string[] _preprocessorSymbols = ["NET8_0"];
#endif

        private static (ImmutableCompilationResult Result, ImmutableArray<Diagnostic> GeneratorDiagnostics) RunGenerator(string source)
        {
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics, string _) = RunGeneratorWithSource(source);
            return (result, diagnostics);
        }

        // Same as RunGenerator, but also returns the generator's own emitted source (every AddSource'd
        // tree beyond the original one, concatenated) — for tests that need to inspect what was actually
        // generated rather than just whether it compiles.
        private static (ImmutableCompilationResult Result, ImmutableArray<Diagnostic> GeneratorDiagnostics, string GeneratedSource) RunGeneratorWithSource(string source)
        {
            var parseOptions = new CSharpParseOptions(preprocessorSymbols: _preprocessorSymbols);
            SyntaxTree tree = CSharpSyntaxTree.ParseText(source, parseOptions);
            MetadataReference[] references = [.. AppDomain.CurrentDomain.GetAssemblies()
                .Where(static a => !a.IsDynamic && !string.IsNullOrEmpty(a.Location))
                .Select(static a => (MetadataReference)MetadataReference.CreateFromFile(a.Location))];
            var compilation = CSharpCompilation.Create(
                "ExcelReader.Generator.Tests.Compilation",
                [tree],
                references,
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

            // parseOptions here is what re-parses the generator's *own* AddSource output (the Guid
            // #if NET9_0_OR_GREATER split) — without passing it explicitly, the driver falls back to
            // default parse options (no symbols defined) for generated sources, regardless of what the
            // original tree above used, and the #else branch would always "win".
            GeneratorDriver driver = CSharpGeneratorDriver.Create([new ExcelRowMapGenerator().AsSourceGenerator()], parseOptions: parseOptions);
            _ = driver.RunGeneratorsAndUpdateCompilation(compilation, out Compilation outputCompilation, out ImmutableArray<Diagnostic> generatorDiagnostics);
            string generatedSource = string.Join("\n----\n", outputCompilation.SyntaxTrees.Skip(1).Select(static t => t.ToString()));
            return (new ImmutableCompilationResult((CSharpCompilation)outputCompilation), generatorDiagnostics, generatedSource);
        }

        // Wraps the updated Compilation so tests can Emit it without repeating the boilerplate.
        private readonly struct ImmutableCompilationResult
        {
            private readonly CSharpCompilation _compilation;

            internal ImmutableCompilationResult(CSharpCompilation compilation)
            {
                _compilation = compilation;
            }

            internal EmitResult Emit()
            {
                using var ms = new MemoryStream();
                return _compilation.Emit(ms);
            }

            internal (EmitResult Emit, byte[] AssemblyBytes) EmitToBytes()
            {
                using var ms = new MemoryStream();
                EmitResult emit = _compilation.Emit(ms);
                return (emit, ms.ToArray());
            }
        }

        [Fact]
        public void NonPartialTypeReportsEXR001()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.NotPartial
                {
                    [ExcelSerializable]
                    public class Model
                    {
                        public string Name { get; set; } = "";
                    }
                }
                """;
            (_, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR001", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void NonPartialContainingTypeReportsEXR002()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.NonPartialContainer
                {
                    public class Outer
                    {
                        [ExcelSerializable]
                        public partial class Model
                        {
                            public string Name { get; set; } = "";
                        }
                    }
                }
                """;
            (_, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR002", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void PartialTypeWithSupportedPropertiesEmitsNoDiagnostics()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.Happy
                {
                    [ExcelSerializable]
                    public partial class Model
                    {
                        public string Name { get; set; } = "";
                        public int Age { get; set; }
                        [ExcelIgnore]
                        public string Ignored { get; set; } = "";
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }

        [Fact]
        public void UnsupportedPropertyTypeIsSkippedWithoutDiagnostics()
        {
            // A type with no built-in reader and no [ExcelConverter] is silently left unmapped, same as
            // the reflection path's "no parser found" outcome for an unrecognized type — no diagnostic
            // unless the property is also [ExcelRequired] (EXR003, tested separately below).
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.Unsupported
                {
                    public struct Point
                    {
                        public int X;
                        public int Y;
                    }

                    [ExcelSerializable]
                    public partial class Model
                    {
                        public string Name { get; set; } = "";
                        public Point Location { get; set; }
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }

        [Fact]
        public void RequiredUnsupportedPropertyReportsEXR003()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.RequiredUnsupported
                {
                    public struct Point
                    {
                        public int X;
                        public int Y;
                    }

                    [ExcelSerializable]
                    public partial class Model
                    {
                        [ExcelRequired]
                        public Point Location { get; set; }
                    }
                }
                """;
            (_, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR003", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void InvalidConverterTypeReportsEXR004()
        {
            const string source = """
                using ExcelReader.Core.Parser;
                using ExcelReader.Core.ValueObjects;

                namespace GeneratorTests.BadConverter
                {
                    public sealed class WrongConverter : IExcelCellConverter<int>
                    {
                        public bool TryConvert(in Cell cell, bool isDate1904, System.IFormatProvider provider, out int value)
                        {
                            value = 0;
                            return true;
                        }
                    }

                    [ExcelSerializable]
                    public partial class Model
                    {
                        // WrongConverter converts int, not string — must not match this property's exact type.
                        [ExcelConverter(typeof(WrongConverter))]
                        public string Name { get; set; } = "";
                    }
                }
                """;
            (_, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR004", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void ValidConverterEmitsNoDiagnosticsAndCompiles()
        {
            const string source = """
                using ExcelReader.Core.Parser;
                using ExcelReader.Core.ValueObjects;
                using ExcelReader.Core.Writer;

                namespace GeneratorTests.GoodConverter
                {
                    public sealed class UpperCaseConverter : IExcelCellConverter<string>, IExcelCellWriter<string>
                    {
                        public bool TryConvert(in Cell cell, bool isDate1904, System.IFormatProvider provider, out string value)
                        {
                            value = cell.GetString().ToUpperInvariant();
                            return true;
                        }

                        public void Write(IRowWriter row, string value)
                        {
                            row.Write(value);
                        }
                    }

                    [ExcelSerializable]
                    public partial class Model
                    {
                        [ExcelConverter(typeof(UpperCaseConverter))]
                        public string Name { get; set; } = "";
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }

        [Fact]
        public void DuplicateHeaderNameReportsEXR006ButStillCompiles()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.DuplicateHeader
                {
                    [ExcelSerializable]
                    public partial class Model
                    {
                        [ExcelColumn("Name")]
                        public string First { get; set; } = "";
                        [ExcelColumn("Name")]
                        public string Second { get; set; } = "";
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR006", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Warning, diagnostic.Severity);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }

        [Fact]
        public void EveryBuiltInTypeAndItsNullableFormCompiles()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.EveryType
                {
                    public enum Kind { Alpha, Beta }

                    [ExcelSerializable]
                    public partial class Model
                    {
                        public string Name { get; set; } = "";
                        public bool Active { get; set; }
                        public System.DateTime BirthDate { get; set; }
                        public System.DateOnly BirthDay { get; set; }
                        public System.TimeOnly BirthTime { get; set; }
                        public System.Guid Id { get; set; }
                        public byte U8 { get; set; }
                        public sbyte I8 { get; set; }
                        public short I16 { get; set; }
                        public ushort U16 { get; set; }
                        public int I32 { get; set; }
                        public uint U32 { get; set; }
                        public long I64 { get; set; }
                        public ulong U64 { get; set; }
                        public float F32 { get; set; }
                        public double F64 { get; set; }
                        public decimal Money { get; set; }
                        public char Letter { get; set; }
                        public System.TimeSpan Duration { get; set; }
                        public System.DateTimeOffset Offset { get; set; }
                        public Kind Category { get; set; }
                        public bool? ActiveN { get; set; }
                        public System.DateTime? BirthDateN { get; set; }
                        public int? I32N { get; set; }
                        public decimal? MoneyN { get; set; }
                        public Kind? CategoryN { get; set; }
                        public System.Guid? IdN { get; set; }
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }

        // Behavioral parity: compiles a self-contained model + driver method into a real in-memory
        // assembly (so ExcelMappedParser<Model> resolves with ordinary compile-time generics inside
        // that assembly — no MakeGenericType reflection gymnastics needed here), loads it, and calls
        // the driver via plain reflection (no generics needed for that part, since the driver method
        // itself is non-generic and returns only strings).
        [Fact]
        public async Task GeneratedMapRoundTripsThroughRealXlsxReaderAndWriter()
        {
            const string source = """
                using System;
                using System.Collections.Generic;
                using System.Globalization;
                using System.IO;
                using System.Threading.Tasks;
                using ExcelReader.Core.Parser;
                using ExcelReader.Core.Reader;
                using ExcelReader.Core.Writer;

                namespace GeneratorTests.RoundTrip
                {
                    public enum Kind { Alpha, Beta }

                    [ExcelSerializable]
                    public partial class Model
                    {
                        public string Name { get; set; } = "";
                        public bool Active { get; set; }
                        public int Age { get; set; }
                        public decimal Balance { get; set; }
                        public Kind Category { get; set; }
                        public int? OptionalAge { get; set; }
                        public Kind? OptionalCategory { get; set; }
                    }

                    public static class TestRunner
                    {
                        // Writes the header/data row directly through the public XlsxRowWriter API
                        // (mirroring what a hand-rolled test fixture would do) rather than through
                        // ExcelRecordMapBuilder<T>.Headers()/WriteRow(), which are internal to
                        // ExcelReader.Core — this synthetic assembly has no InternalsVisibleTo grant for
                        // them. The read side below still exercises the real generated
                        // IExcelRowMap<Model>/ExcelMappedParser<Model> path end to end.
                        public static async Task<string[]> RunAsync()
                        {
                            var writeStream = new MemoryStream();
                            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(writeStream, leaveOpen: true))
                            {
                                await wb.StartAsync();
                                XlsxSheetWriter sheet = wb.AddSheet("S1");
                                await sheet.StartAsync();
                                XlsxRowWriter header = await sheet.StartRowAsync();
                                await using (header.ConfigureAwait(false))
                                {
                                    header.Write("Name");
                                    header.Write("Active");
                                    header.Write("Age");
                                    header.Write("Balance");
                                    header.Write("Category");
                                    header.Write("OptionalAge");
                                    header.Write("OptionalCategory");
                                }
                                XlsxRowWriter row = await sheet.StartRowAsync();
                                await using (row.ConfigureAwait(false))
                                {
                                    row.Write("Alice");
                                    row.Write(true);
                                    row.Write(30);
                                    row.Write(12.5m);
                                    row.Write("Beta");
                                    row.Write(7);
                                    row.Write("Alpha");
                                }
                                await sheet.EndAsync();
                                await wb.EndAsync();
                            }

                            writeStream.Position = 0;
                            await using XlsxReader reader = await Excel.FromAsync(writeStream);
                            var results = new List<Model>();
                            foreach (Model m in new ExcelMappedParser<Model>().Parse(reader))
                            {
                                results.Add(m);
                            }
                            Model m0 = results[0];
                            return new[]
                            {
                                results.Count.ToString(CultureInfo.InvariantCulture),
                                m0.Name,
                                m0.Active.ToString(CultureInfo.InvariantCulture),
                                m0.Age.ToString(CultureInfo.InvariantCulture),
                                m0.Balance.ToString(CultureInfo.InvariantCulture),
                                m0.Category.ToString(),
                                m0.OptionalAge.ToString(),
                                m0.OptionalCategory.ToString(),
                            };
                        }
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Assert.Empty(diagnostics);
            (EmitResult Emit, byte[] AssemblyBytes) emitted = result.EmitToBytes();
            Assert.True(emitted.Emit.Success, string.Join(Environment.NewLine, emitted.Emit.Diagnostics.Select(static d => d.ToString())));

            Assembly assembly = Assembly.Load(emitted.AssemblyBytes);
            Type runnerType = assembly.GetType("GeneratorTests.RoundTrip.TestRunner")!;
            MethodInfo method = runnerType.GetMethod("RunAsync")!;
            var task = (Task<string[]>)method.Invoke(null, null)!;
            string[] values = await task;

            Assert.Equal("1", values[0]);
            Assert.Equal("Alice", values[1]);
            Assert.Equal("True", values[2]);
            Assert.Equal("30", values[3]);
            Assert.Equal("12.5", values[4]);
            Assert.Equal("Beta", values[5]);
            Assert.Equal("7", values[6]);
            Assert.Equal("Alpha", values[7]);
        }

        [Fact]
        public void GenericTypeReportsEXR007()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.Generic
                {
                    [ExcelSerializable]
                    public partial class Box<TItem>
                    {
                        public string Name { get; set; } = "";
                    }
                }
                """;
            (_, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR007", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void GenericContainingTypeReportsEXR007()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.GenericContainer
                {
                    public partial class Outer<TItem>
                    {
                        [ExcelSerializable]
                        public partial class Model
                        {
                            public string Name { get; set; } = "";
                        }
                    }
                }
                """;
            (_, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR007", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void EmptyTypeReportsEXR005AndStillCompiles()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.Empty
                {
                    [ExcelSerializable]
                    public partial class Model
                    {
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR005", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Warning, diagnostic.Severity);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }

        [Fact]
        public void WriteOnlyPropertyStillCompiles()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.WriteOnly
                {
                    [ExcelSerializable]
                    public partial class Model
                    {
                        private string _name = "";
                        public string Name { set => _name = value; }
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics, string generated) = RunGeneratorWithSource(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
            Assert.Contains(".Property([\"Name\"]", generated, StringComparison.Ordinal);
        }

        [Fact]
        public void RecordClassWithSettablePropertiesCompiles()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.RecordClass
                {
                    [ExcelSerializable]
                    public partial record class Rec
                    {
                        public string Name { get; set; } = "";
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }

        [Fact]
        public void RecordStructWithSettablePropertiesCompiles()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.RecordStruct
                {
                    [ExcelSerializable]
                    public partial record struct Rec
                    {
                        public string Name { get; set; }
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }

        [Fact]
        public void InitOnlyPropertyIsNotMappedForRead()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.InitOnly
                {
                    [ExcelSerializable]
                    public partial class Model
                    {
                        public string Name { get; init; } = "";
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics, string generated) = RunGeneratorWithSource(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
            Assert.DoesNotContain(".Property([\"Name\"]", generated, StringComparison.Ordinal);
            Assert.Contains(".Column(\"Name\"", generated, StringComparison.Ordinal);
        }

        [Fact]
        public void PositionalRecordReportsEXR008()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.Positional
                {
                    [ExcelSerializable]
                    public partial record Foo(string A);
                }
                """;
            (_, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Diagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal("EXR008", diagnostic.Id);
            Assert.Equal(DiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void InheritedPropertiesAreMapped()
        {
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.Inherited
                {
                    public class Base
                    {
                        public string Inherited { get; set; } = "";
                    }

                    [ExcelSerializable]
                    public partial class Derived : Base
                    {
                        public int Own { get; set; }
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics, string generated) = RunGeneratorWithSource(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
            Assert.Contains(".Property([\"Own\"]", generated, StringComparison.Ordinal);
            Assert.Contains(".Property([\"Inherited\"]", generated, StringComparison.Ordinal);
            int ownIndex = generated.IndexOf(".Property([\"Own\"]", StringComparison.Ordinal);
            int inheritedIndex = generated.IndexOf(".Property([\"Inherited\"]", StringComparison.Ordinal);
            Assert.True(ownIndex < inheritedIndex, "Declared-on-type properties must be emitted before inherited ones.");
        }
    }
}
