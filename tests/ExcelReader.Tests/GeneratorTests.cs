using System.Collections.Immutable;
using System.Linq;
using ExcelReader.Generator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;

namespace ExcelReader.Tests
{
    // Feature A2 (docs/v2-plan.md §2.4): the generator's own plumbing — triggers on [ExcelSerializable],
    // requires 'partial' up the containing-type chain (EXR001/EXR002), and emits code that type-checks
    // against the real public ExcelRowMapBuilder<T>/ExcelRecordMapBuilder<T>/ExcelCellReaders surface.
    // Runs ExcelRowMapGenerator directly via CSharpGeneratorDriver over an in-memory Compilation, rather
    // than through the SDK's live build-time analyzer pipeline — a deliberately broken model (missing
    // 'partial', for the EXR001/EXR002 cases) would otherwise fail this whole project's own build.
    public class GeneratorTests
    {
        private static (ImmutableCompilationResult Result, ImmutableArray<Diagnostic> GeneratorDiagnostics) RunGenerator(string source)
        {
            SyntaxTree tree = CSharpSyntaxTree.ParseText(source);
            MetadataReference[] references = [.. AppDomain.CurrentDomain.GetAssemblies()
                .Where(static a => !a.IsDynamic && !string.IsNullOrEmpty(a.Location))
                .Select(static a => (MetadataReference)MetadataReference.CreateFromFile(a.Location))];
            var compilation = CSharpCompilation.Create(
                "ExcelReader.Generator.Tests.Compilation",
                [tree],
                references,
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

            GeneratorDriver driver = CSharpGeneratorDriver.Create(new ExcelRowMapGenerator());
            _ = driver.RunGeneratorsAndUpdateCompilation(compilation, out Compilation outputCompilation, out ImmutableArray<Diagnostic> generatorDiagnostics);
            return (new ImmutableCompilationResult((CSharpCompilation)outputCompilation), generatorDiagnostics);
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
            // A2 scope: only string/int are supported. A property of any other type is silently left
            // unmapped, same as the reflection path's "no parser found" outcome for an unrecognized type
            // — no diagnostic yet (that upgrade, EXR003 for [ExcelRequired] on an unsupported type,
            // belongs to a later phase once more types are supported).
            const string source = """
                using ExcelReader.Core.Parser;

                namespace GeneratorTests.Unsupported
                {
                    [ExcelSerializable]
                    public partial class Model
                    {
                        public string Name { get; set; } = "";
                        public System.DateTime BirthDate { get; set; }
                    }
                }
                """;
            (ImmutableCompilationResult result, ImmutableArray<Diagnostic> diagnostics) = RunGenerator(source);
            Assert.Empty(diagnostics);
            EmitResult emit = result.Emit();
            Assert.True(emit.Success, string.Join(Environment.NewLine, emit.Diagnostics.Select(static d => d.ToString())));
        }
    }
}
