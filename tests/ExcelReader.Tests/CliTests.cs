using ExcelReader.Cli;

namespace ExcelReader.Tests
{
    public sealed class CliTests
    {
        private static string Fixture(string name) => Path.Combine("data", name);

        private static (int Code, string Out, string Err) Sheets(string path)
        {
            StringWriter stdout = new();
            StringWriter stderr = new();
            int code = CliCommands.Sheets(path, stdout, stderr);
            return (code, stdout.ToString(), stderr.ToString());
        }

        [Fact]
        public void Should_ListEverySheet_When_RunningSheets()
        {
            (int code, string output, string error) = Sheets(Fixture("RealExcel.xlsb"));

            Assert.Equal(0, code);
            Assert.Empty(error);

            // One line per sheet, "<index>\t<name>".
            string[] lines = output.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            Assert.NotEmpty(lines);
            Assert.StartsWith("0\t", lines[0], StringComparison.Ordinal);
        }

        [Fact]
        public void Should_ReturnOneAndWriteToStderr_When_TheFileDoesNotExist()
        {
            (int code, string output, string error) = Sheets(
                Path.Combine(Path.GetTempPath(), "no-such-workbook.xlsx"));

            Assert.Equal(1, code);
            Assert.Empty(output);
            Assert.NotEmpty(error);
            // A message the user can act on, not a stack trace.
            Assert.DoesNotContain("   at ", error, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_SelectBySheetName_When_SheetIsNotNumeric()
        {
            using ExcelReader.Core.Reader.IExcelRowReader byIndex = CliCommands.Open(Fixture("RealExcel.xlsb"), "0");
            string firstSheetName = byIndex.SheetName;

            using ExcelReader.Core.Reader.IExcelRowReader byName = CliCommands.Open(Fixture("RealExcel.xlsb"), firstSheetName);
            Assert.Equal(firstSheetName, byName.SheetName);
        }

        [Fact]
        public void Should_Throw_When_TheNamedSheetIsAbsent()
        {
            ArgumentException error = Assert.Throws<ArgumentException>(
                () => CliCommands.Open(Fixture("RealExcel.xlsb"), "NoSuchSheet"));

            Assert.Contains("NoSuchSheet", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_DefaultToTheFirstSheet_When_SheetIsNull()
        {
            using ExcelReader.Core.Reader.IExcelRowReader reader = CliCommands.Open(Fixture("RealExcel.xlsb"), null);

            Assert.Equal(reader.SheetNameAt(0), reader.SheetName);
        }

        private static (int Code, string Out, string Err) Convert(
            string path, string? sheet = null, string? output = null, char delimiter = ',')
        {
            using MemoryStream stdout = new();
            StringWriter stderr = new();
            int code = CliCommands.Convert(path, sheet, output, delimiter, stdout, stderr);
            return (code, System.Text.Encoding.UTF8.GetString(stdout.ToArray()), stderr.ToString());
        }

        [Fact]
        public void Should_WriteCsvToTheOutputStream_When_NoOutputFileIsGiven()
        {
            (int code, string output, string error) = Convert(Fixture("RealExcel.xlsb"));

            Assert.Equal(0, code);
            Assert.Empty(error);
            Assert.NotEmpty(output);
            Assert.Contains(",", output, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_WriteCsvToAFile_When_OutputIsGiven()
        {
            string target = Path.Combine(Path.GetTempPath(), $"excelreader-cli-{Guid.NewGuid():N}.csv");
            try
            {
                (int code, string output, string error) = Convert(Fixture("RealExcel.xlsb"), output: target);

                Assert.Equal(0, code);
                Assert.Empty(error);
                // Nothing on the stdout stream: the CSV went to the file.
                Assert.Empty(output);
                Assert.True(File.Exists(target));
                Assert.NotEmpty(File.ReadAllText(target));
            }
            finally
            {
                File.Delete(target);
            }
        }

        [Fact]
        public void Should_UseTheGivenDelimiter_When_DelimiterIsOverridden()
        {
            (int code, string output, _) = Convert(Fixture("RealExcel.xlsb"), delimiter: ';');

            Assert.Equal(0, code);
            Assert.Contains(";", output, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_ReturnOne_When_TheNamedSheetIsAbsent()
        {
            (int code, _, string error) = Convert(Fixture("RealExcel.xlsb"), sheet: "NoSuchSheet");

            Assert.Equal(1, code);
            Assert.Contains("NoSuchSheet", error, StringComparison.Ordinal);
        }

        private static (int Code, string Out, string Err) Schema(
            string path, string? sheet = null, int headerRow = 1, int sampleSize = 100)
        {
            StringWriter stdout = new();
            StringWriter stderr = new();
            int code = CliCommands.Schema(path, sheet, headerRow, sampleSize, stdout, stderr);
            return (code, stdout.ToString(), stderr.ToString());
        }

        [Fact]
        public void Should_PrintOneLinePerColumn_When_RunningSchema()
        {
            (int code, string output, string error) = Schema(Fixture("RealExcel.xlsb"));

            Assert.Equal(0, code);
            Assert.Empty(error);

            string[] lines = output.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            Assert.NotEmpty(lines);
            // "<index>\t<name>\t<type>[?]"
            foreach (string line in lines)
            {
                string[] parts = line.TrimEnd('\r').Split('\t');
                Assert.Equal(3, parts.Length);
                Assert.True(int.TryParse(parts[0], System.Globalization.CultureInfo.InvariantCulture, out _));
                Assert.NotEmpty(parts[2]);
            }
        }

        [Fact]
        public void Should_PrintAnEmptyNameField_When_HeaderRowIsZero()
        {
            (int code, string output, _) = Schema(Fixture("RealExcel.xlsb"), headerRow: 0);

            Assert.Equal(0, code);
            // A null name renders as an empty middle field, never as the literal "null".
            string firstLine = output.Split('\n', StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd('\r');
            Assert.Equal(string.Empty, firstLine.Split('\t')[1]);
        }

        [Fact]
        public void Should_MarkNullableColumnsWithAQuestionMark_When_PrintingTheSchema()
        {
            // A CSV with a gap guarantees at least one nullable column, independent of the fixture.
            string target = Path.Combine(Path.GetTempPath(), $"excelreader-cli-{Guid.NewGuid():N}.csv");
            File.WriteAllText(target, "Id,Note\n1,here\n2,\n");
            try
            {
                (int code, string output, _) = Schema(target);

                Assert.Equal(0, code);
                Assert.Contains("?", output, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(target);
            }
        }
    }
}
