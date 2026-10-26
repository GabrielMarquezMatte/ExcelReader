using ExcelReader.Cli;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Xlsx;
using ExcelReader.Tests.Crypto;

namespace ExcelReader.Tests.Cli
{
    public sealed class CliTests : IDisposable
    {
        private readonly List<string> _temporaryFiles = [];

        public void Dispose()
        {
            foreach (string file in _temporaryFiles)
            {
                File.Delete(file);
            }
        }

        private string MultiSheetPath()
        {
            string path = Path.Combine(Path.GetTempPath(), $"cli-multi-sheet-{Guid.NewGuid():N}.xlsx");
            _temporaryFiles.Add(path);
            using FileStream file = File.Create(path);
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(file, leaveOpen: true);
            foreach (string name in (string[])["First", "Second", "Third"])
            {
                using XlsxSheetWriter sheet = workbook.AddSheet(name);
                sheet.End();
            }
            workbook.End();
            return path;
        }

        private static string Fixture(string name)
        {
            return Path.Combine("data", name);
        }

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

            using IExcelWorkbook reader = Excel.Open(Fixture("RealExcel.xlsb"));

            string[] lines = output.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            Assert.Equal(reader.SheetCount, lines.Length);
            for (int i = 0; i < lines.Length; i++)
            {
                string[] parts = lines[i].TrimEnd('\r').Split('\t');
                Assert.Equal(2, parts.Length);
                Assert.Equal(i.ToString(System.Globalization.CultureInfo.InvariantCulture), parts[0]);
                Assert.Equal(reader.SheetAt(i).Name, parts[1]);
            }
        }

        [Fact]
        public void Should_ReturnOneAndWriteToStderr_When_TheFileDoesNotExist()
        {
            (int code, string output, string error) = Sheets(
                Path.Combine(Path.GetTempPath(), "no-such-workbook.xlsx"));

            Assert.Equal(1, code);
            Assert.Empty(output);
            Assert.NotEmpty(error);
            Assert.DoesNotContain("   at ", error, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_SelectBySheetName_When_SheetIsNotNumeric()
        {
            using IExcelWorkbook byIndex = CliCommands.Open(Fixture("RealExcel.xlsb"), "0", out IExcelSheet byIndexSheet);
            string firstSheetName = byIndexSheet.Name;

            using IExcelWorkbook byName = CliCommands.Open(Fixture("RealExcel.xlsb"), firstSheetName, out IExcelSheet byNameSheet);
            Assert.Equal(firstSheetName, byNameSheet.Name);
        }

        [Fact]
        public void Should_PreferASheetNamedLikeANumber_Over_TheSheetAtThatIndex()
        {
            string path = Path.Combine(Path.GetTempPath(), $"cli-numeric-sheet-{Guid.NewGuid():N}.xlsx");
            try
            {
                using (FileStream file = File.Create(path))
                using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(file, leaveOpen: true))
                {
                    foreach (string name in (string[])["Data", "0"])
                    {
                        using XlsxSheetWriter sheet = workbook.AddSheet(name);
                        sheet.End();
                    }
                    workbook.End();
                }

                using IExcelWorkbook named = CliCommands.Open(path, "0", out IExcelSheet namedSheet);
                Assert.Equal("0", namedSheet.Name);

                using IExcelWorkbook indexed = CliCommands.Open(path, "1", out IExcelSheet indexedSheet);
                Assert.Equal("0", indexedSheet.Name);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void Should_Throw_When_TheNamedSheetIsAbsent()
        {
            ArgumentException error = Assert.Throws<ArgumentException>(
                () => CliCommands.Open(Fixture("RealExcel.xlsb"), "NoSuchSheet", out _));

            Assert.Contains("NoSuchSheet", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_DefaultToTheFirstSheet_When_SheetIsNull()
        {
            using IExcelWorkbook reader = CliCommands.Open(Fixture("RealExcel.xlsb"), null, out IExcelSheet selected);

            Assert.Equal(reader.SheetAt(0).Name, selected.Name);
        }

        [Fact]
        public void Should_FailWithAMessage_When_TheSheetIndexIsOutOfRange()
        {
            using StringWriter stdout = new();
            using StringWriter stderr = new();
            int exit = CliCommands.Schema(MultiSheetPath(), sheet: "7", headerRow: 1, sampleSize: 10, stdout, stderr);
            Assert.Equal(1, exit);
            Assert.NotEmpty(stderr.ToString());
            Assert.Empty(stdout.ToString());
        }

        [Fact]
        public void Should_ReturnTheWorkbookAndTheSelectedSheet_When_Opening()
        {
            using IExcelWorkbook workbook = CliCommands.Open(MultiSheetPath(), sheet: "1", out IExcelSheet selected);
            Assert.Equal(1, selected.Index);
            Assert.Equal(workbook.SheetAt(1).Name, selected.Name);

            using IExcelWorkbook first = CliCommands.Open(MultiSheetPath(), sheet: null, out IExcelSheet defaulted);
            Assert.Equal(0, defaulted.Index);
        }

        [Fact]
        public void Should_OpenAnEncryptedWorkbook_When_PasswordIsGiven()
        {
            using IExcelWorkbook reader = CliCommands.Open(
                Path.Combine("data", "encrypted", "agile-aes256-sha512.xlsx"), null, out _, EncryptedFixtures.Password);

            Assert.True(reader.SheetCount > 0);
        }

        [Fact]
        public void Should_ReturnOneAndWriteToStderr_When_AnEncryptedWorkbookHasNoPassword()
        {
            (int code, string output, string error) = Sheets(
                Path.Combine("data", "encrypted", "agile-aes256-sha512.xlsx"));

            Assert.Equal(1, code);
            Assert.Empty(output);
            Assert.NotEmpty(error);
        }

        private static (int Code, string Out, string Err) Convert(
            string path, string? sheet = null, string? output = null, string? format = null, char delimiter = ',')
        {
            using MemoryStream stdout = new();
            StringWriter stderr = new();
            int code = CliCommands.Convert(path, sheet, output, format, delimiter, stdout, stderr);
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
            (int commaCode, string commaOutput, _) = Convert(Fixture("RealExcel.xlsb"));
            (int code, string output, _) = Convert(Fixture("RealExcel.xlsb"), delimiter: ';');

            Assert.Equal(0, commaCode);
            Assert.Equal(0, code);

            string commaHeader = commaOutput.Split('\n', StringSplitOptions.RemoveEmptyEntries)[0];
            string semicolonHeader = output.Split('\n', StringSplitOptions.RemoveEmptyEntries)[0];
            int fieldCount = commaHeader.Split(',').Length;
            Assert.True(fieldCount > 1, "fixture's header row should have more than one field");
            Assert.Equal(fieldCount, semicolonHeader.Split(';').Length);
            Assert.DoesNotContain(",", semicolonHeader, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_ReturnOne_When_TheNamedSheetIsAbsent()
        {
            (int code, _, string error) = Convert(Fixture("RealExcel.xlsb"), sheet: "NoSuchSheet");

            Assert.Equal(1, code);
            Assert.Contains("NoSuchSheet", error, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        [InlineData("xls")]
        public void Should_ConvertToAnotherSpreadsheetFormat_When_OutputExtensionNamesOne(string extension)
        {
            string target = Path.Combine(Path.GetTempPath(), $"excelreader-cli-{Guid.NewGuid():N}.{extension}");
            try
            {
                (int code, string output, string error) = Convert(Fixture("RealExcel.xlsb"), output: target);

                Assert.Equal(0, code);
                Assert.Empty(error);
                Assert.Empty(output);

                using IExcelWorkbook source = Excel.Open(Fixture("RealExcel.xlsb"));
                using IExcelRowEnumerator sourceRows = source.FirstSheet.GetEnumerator();
                using IExcelWorkbook written = Excel.Open(target);
                using IExcelRowEnumerator writtenRows = written.FirstSheet.GetEnumerator();

                int rowCount = 0;
                while (sourceRows.MoveNext())
                {
                    Assert.True(writtenRows.MoveNext(), $"written file ran out of rows at row {rowCount}");
                    Core.Reader.Row sourceRow = sourceRows.Current;
                    Core.Reader.Row writtenRow = writtenRows.Current;
                    Assert.Equal(sourceRow.ColumnCount, writtenRow.ColumnCount);
                    for (int column = 0; column < sourceRow.ColumnCount; column++)
                    {
                        Assert.Equal(sourceRow[column].GetString(), writtenRow[column].GetString());
                    }
                    rowCount++;
                }
                Assert.False(writtenRows.MoveNext(), "written file has more rows than the source");
                Assert.True(rowCount > 1, "fixture should have more than a header row");
            }
            finally
            {
                File.Delete(target);
            }
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        [InlineData("xls")]
        public void Should_KeepDates_When_ConvertingA1904Workbook(string extension)
        {
            var date = new DateTime(2024, 5, 6, 12, 30, 0, DateTimeKind.Unspecified);
            string source = Path.Combine(Path.GetTempPath(), $"excelreader-cli-{Guid.NewGuid():N}.xlsx");
            string target = Path.ChangeExtension(source, "out." + extension);
            try
            {
                using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(File.Create(source), options: new XlsxWriterOptions { Date1904 = true }))
                {
                    XlsxSheetWriter sheet = wb.AddSheet("S");
                    using (XlsxRowWriter row = sheet.StartRow())
                    {
                        row.Write(date);
                    }
                    sheet.End();
                    wb.End();
                }

                (int code, _, string error) = Convert(source, output: target);

                Assert.Equal(0, code);
                Assert.Empty(error);
                using IExcelWorkbook written = Excel.Open(target);
                using IExcelRowEnumerator rows = written.FirstSheet.GetEnumerator();
                Assert.True(rows.MoveNext());
                Assert.True(rows.Current[0].TryGetDateTime(written.IsDate1904, out DateTime readBack));
                Assert.Equal(date, readBack);
            }
            finally
            {
                File.Delete(source);
                File.Delete(target);
            }
        }

        [Fact]
        public void Should_UseTheExplicitFormat_When_FormatOverridesTheOutputExtension()
        {
            (int code, string output, string error) = Convert(Fixture("RealExcel.xlsb"), format: "xlsx");

            Assert.Equal(0, code);
            Assert.Empty(error);
            Assert.StartsWith("PK", output, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_ReturnOneAndListValidExtensions_When_TheOutputExtensionIsUnrecognized()
        {
            string target = Path.Combine(Path.GetTempPath(), $"excelreader-cli-{Guid.NewGuid():N}.txt");

            (int code, string output, string error) = Convert(Fixture("RealExcel.xlsb"), output: target);

            Assert.Equal(1, code);
            Assert.Empty(output);
            Assert.Contains(".txt", error, StringComparison.Ordinal);
            Assert.Contains(".xlsx", error, StringComparison.Ordinal);
            Assert.False(File.Exists(target));
        }

        [Fact]
        public void Should_ReturnOne_When_TheFormatNameIsUnknown()
        {
            (int code, _, string error) = Convert(Fixture("RealExcel.xlsb"), format: "pdf");

            Assert.Equal(1, code);
            Assert.Contains("pdf", error, StringComparison.Ordinal);
        }

        private string TemporaryPath(string extension)
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-cli-{Guid.NewGuid():N}.{extension}");
            _temporaryFiles.Add(path);
            return path;
        }

        private string WriteXlsx(Action<XlsxRowWriter> writeRow)
        {
            string path = TemporaryPath("xlsx");
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(File.Create(path));
            using XlsxSheetWriter sheet = workbook.AddSheet("Data");
            using (XlsxRowWriter row = sheet.StartRow())
            {
                writeRow(row);
            }
            sheet.End();
            workbook.End();
            return path;
        }

        private string WriteRawXlsx(string rows)
        {
            string path = TemporaryPath("xlsx");
            using MemoryStream built = WorkbookBuilder.Build(rows);
            File.WriteAllBytes(path, built.ToArray());
            return path;
        }

        private static string FirstCsvLine(string csv)
        {
            return csv.Split('\n')[0].TrimEnd('\r');
        }

        [Fact]
        public void Should_KeepTheExistingOutput_When_ConversionFails()
        {
            string directory = Path.Combine(Path.GetTempPath(), $"excelreader-cli-{Guid.NewGuid():N}");
            Directory.CreateDirectory(directory);
            try
            {
                string target = Path.Combine(directory, "out.xlsx");
                File.WriteAllText(target, "ORIGINAL");
                string source = WriteRawXlsx("""<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2" t="b"><v>maybe</v></c></row>""");

                (int code, _, string error) = Convert(source, output: target);

                Assert.Equal(1, code);
                Assert.NotEmpty(error);
                Assert.Equal("ORIGINAL", File.ReadAllText(target));
                Assert.Equal([target], Directory.GetFiles(directory));
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void Should_ReplaceTheExistingOutput_When_ConversionSucceeds()
        {
            string target = TemporaryPath("csv");
            File.WriteAllText(target, "ORIGINAL");

            (int code, _, string error) = Convert(WriteXlsx(row => row.Write("new")), output: target);

            Assert.Equal(0, code);
            Assert.Empty(error);
            Assert.Equal("new", FirstCsvLine(File.ReadAllText(target)));
        }

        [Fact]
        public void Should_KeepColumnGaps_When_ConvertingSparseRowsToCsv()
        {
            string source = WriteXlsx(row =>
            {
                row.Skip();
                row.Write("B");
                row.Skip();
                row.Write("D");
            });

            (int code, string output, string error) = Convert(source);

            Assert.Equal(0, code);
            Assert.Empty(error);
            Assert.Equal(",B,,D", FirstCsvLine(output));
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        [InlineData("xls")]
        public void Should_KeepColumnGaps_When_ConvertingSparseRowsToASpreadsheet(string extension)
        {
            string source = WriteXlsx(row =>
            {
                row.Write("A");
                row.Skip();
                row.Write("C");
            });
            string target = TemporaryPath(extension);

            (int code, _, string error) = Convert(source, output: target);

            Assert.Equal(0, code);
            Assert.Empty(error);
            using IExcelWorkbook written = Excel.Open(target);
            using IExcelRowEnumerator rows = written.FirstSheet.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal("A", rows.Current[0].GetString());
            Assert.Equal(string.Empty, rows.Current[1].GetString());
            Assert.Equal("C", rows.Current[2].GetString());
        }

        [Fact]
        public void Should_ConvertEveryBooleanSpelling_When_WritingCsv()
        {
            string source = WriteRawXlsx("""<row r="1"><c r="A1" t="b"><v>false</v></c><c r="B1" t="b"><v>true</v></c><c r="C1" t="b"><v>0</v></c><c r="D1" t="b"><v>1</v></c></row>""");

            (int code, string output, string error) = Convert(source);

            Assert.Equal(0, code);
            Assert.Empty(error);
            Assert.Equal("false,true,false,true", FirstCsvLine(output));
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        [InlineData("xls")]
        public void Should_ConvertEveryBooleanSpelling_When_WritingASpreadsheet(string extension)
        {
            string source = WriteRawXlsx("""<row r="1"><c r="A1" t="b"><v>false</v></c><c r="B1" t="b"><v>true</v></c><c r="C1" t="b"><v>0</v></c><c r="D1" t="b"><v>1</v></c></row>""");
            string target = TemporaryPath(extension);

            (int code, _, string error) = Convert(source, output: target);

            Assert.Equal(0, code);
            Assert.Empty(error);
            using IExcelWorkbook written = Excel.Open(target);
            using IExcelRowEnumerator rows = written.FirstSheet.GetEnumerator();
            Assert.True(rows.MoveNext());
            bool[] expected = [false, true, false, true];
            for (int column = 0; column < expected.Length; column++)
            {
                Core.Reader.Cell cell = rows.Current[column];
                Assert.True(ExcelCellReaders.Bool(in cell, default, out bool value));
                Assert.Equal(expected[column], value);
            }
        }

        [Fact]
        public void Should_KeepExactIntegers_When_ConvertingNumbersAbove2Pow53()
        {
            string source = WriteXlsx(row =>
            {
                row.Write(9007199254740993L);
                row.Write(12345678901234567890.5m);
                row.Write(0.1);
                row.Write(1e-30);
            });
            string target = TemporaryPath("xlsx");

            (int code, string output, string error) = Convert(source);
            (int xlsxCode, _, _) = Convert(source, output: target);

            Assert.Equal(0, code);
            Assert.Empty(error);
            Assert.Equal("9007199254740993,12345678901234567890.5,0.1,1E-30", FirstCsvLine(output));
            Assert.Equal(0, xlsxCode);
            using IExcelWorkbook written = Excel.Open(target);
            using IExcelRowEnumerator rows = written.FirstSheet.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.True(rows.Current[0].TryParse(null, out long integer));
            Assert.Equal(9007199254740993L, integer);
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        [InlineData("xls")]
        public void Should_ConvertCsvToASpreadsheet_When_TheSourceSheetHasNoName(string extension)
        {
            string source = TemporaryPath("csv");
            File.WriteAllText(source, "a,b\n1,2\n");
            string target = TemporaryPath(extension);

            (int code, _, string error) = Convert(source, output: target);

            Assert.Equal(0, code);
            Assert.Empty(error);
            using IExcelWorkbook written = Excel.Open(target);
            Assert.Equal("Sheet1", written.FirstSheet.Name);
            using IExcelRowEnumerator rows = written.FirstSheet.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal("a", rows.Current[0].GetString());
            Assert.Equal("b", rows.Current[1].GetString());
            Assert.True(rows.MoveNext());
            Assert.Equal("1", rows.Current[0].GetString());
            Assert.Equal("2", rows.Current[1].GetString());
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
            string firstLine = output.Split('\n', StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd('\r');
            Assert.Equal(string.Empty, firstLine.Split('\t')[1]);
        }

        [Fact]
        public void Should_MarkNullableColumnsWithAQuestionMark_When_PrintingTheSchema()
        {
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
