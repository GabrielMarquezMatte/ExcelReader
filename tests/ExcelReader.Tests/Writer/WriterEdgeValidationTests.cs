using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer
{
    /// <summary>
    /// Inputs the writers used to accept and Excel then refused or altered. Each was confirmed against
    /// Excel 16; these tests pin the writer side.
    /// </summary>
    public sealed class WriterEdgeValidationTests
    {
        public static TheoryData<string> Formats => ["xlsx", "xlsb", "xls"];

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void XlsxEscapesCarriageReturnsAndNonCharacters(bool sharedStrings)
        {
            string[] texts = ["a\rb", "a\r\nb", "\r", "a￾b", "a￿b", "plain"];
            using MemoryStream ms = new();
            using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true, new XlsxWriterOptions { UseSharedStrings = sharedStrings }))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                using (XlsxRowWriter row = sheet.StartRow())
                {
                    foreach (string text in texts)
                    {
                        row.Write(text);
                    }
                    row.WriteUtf8(Encoding.UTF8.GetBytes("u￾v"));
                }
                sheet.End();
                wb.End();
            }

            byte[] xml = Entry(ms.ToArray(), sharedStrings ? "xl/sharedStrings.xml" : "xl/worksheets/sheet1.xml");
            Assert.DoesNotContain((byte)'\r', xml);
            Assert.Equal(-1, xml.AsSpan().IndexOf("￾"u8));
            Assert.Equal(-1, xml.AsSpan().IndexOf("￿"u8));
            Assert.True(xml.AsSpan().IndexOf("a_x000D_b"u8) >= 0);
            Assert.True(xml.AsSpan().IndexOf("a_xFFFE_b"u8) >= 0);
            Assert.True(xml.AsSpan().IndexOf("a_xFFFF_b"u8) >= 0);

            using IExcelWorkbook read = Excel.Open(ms.ToArray());
            using IExcelRowEnumerator rows = read.FirstSheet.GetEnumerator();
            Assert.True(rows.MoveNext());
            for (int i = 0; i < texts.Length; i++)
            {
                Assert.Equal(texts[i], rows.Current[i].GetString());
            }
            Assert.Equal("u￾v", rows.Current[texts.Length].GetString());
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void SheetNamesExcelRefusesAreRejected(string format)
        {
            foreach (string name in new[] { "'a", "a'", "'", "a\u0001b", "a\tb", "a\nb", "a￾b", "a￿b" })
            {
                Assert.Throws<ArgumentException>(() => WithWorkbook(format, wb => wb.AddSheet(name)));
            }
            WithWorkbook(format, wb => wb.AddSheet("it's"));
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void ColumnSettingsBeyondTheFormatAreRejected(string format)
        {
            int columns = string.Equals(format, "xls", StringComparison.Ordinal) ? 256 : 16_384;
            Assert.Throws<ArgumentOutOfRangeException>(() => WithWorkbook(format, wb => wb.Width(columns, 10)));
            Assert.Throws<ArgumentOutOfRangeException>(() => WithWorkbook(format, wb => wb.Style(columns)));
            Assert.Throws<ArgumentOutOfRangeException>(() => WithWorkbook(format, wb => wb.Width(0, 255.5)));
            Assert.Throws<ArgumentOutOfRangeException>(() => WithWorkbook(format, wb => wb.Width(0, double.PositiveInfinity)));
            WithWorkbook(format, wb =>
            {
                wb.Width(columns - 1, 255);
                wb.Style(columns - 1);
            });
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void NumberFormatsStopAtWhatExcelOpens(string format)
        {
            WithWorkbook(format, wb =>
            {
                for (int i = 0; i < 200; i++)
                {
                    wb.AddStyle(new CellStyle { NumberFormat = $"0.00\"u{i}\"" });
                }
                wb.AddStyle(new CellStyle { NumberFormat = "0.00\"u7\"", Bold = true });
                ExcelLimitExceededException error = Assert.Throws<ExcelLimitExceededException>(
                    () => wb.AddStyle(new CellStyle { NumberFormat = "0.000" }));
                Assert.Equal(200, error.Limit);
                wb.AddSheet("S1");
            });
            Assert.Throws<ArgumentException>(
                () => WithWorkbook(format, wb => wb.AddStyle(new CellStyle { NumberFormat = new string('0', 256) })));
            WithWorkbook(format, wb =>
            {
                wb.AddStyle(new CellStyle { NumberFormat = new string('0', 255) });
                wb.AddSheet("S1");
            });
        }

        [Fact]
        public void AValueWithNoNumericFormIsRefusedTheSameWayEverywhere()
        {
            using MemoryStream ms = new();
            using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                using (XlsxRowWriter row = sheet.StartRow())
                {
                    Assert.Throws<ArgumentException>(() => row.Write('c'));
                }
                sheet.End();
            }
            using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                using (XlsbRowWriter row = sheet.StartRow())
                {
                    Assert.Throws<ArgumentException>(() => row.Write('c'));
                }
                sheet.End();
            }
            using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsSheetWriter sheet = wb.AddSheet("S1");
                using (XlsRowWriter row = sheet.StartRow())
                {
                    Assert.Throws<ArgumentException>(() => row.Write('c'));
                }
                sheet.End();
            }
        }

        [Fact]
        public void ALimitMessageNamesNoUnit()
        {
            var error = new ExcelLimitExceededException("Rows", 1_048_576, 1_048_577);

            Assert.Equal("Rows limit exceeded: 1048577 exceeds 1048576.", error.Message);
        }

        private static void WithWorkbook(string format, Action<Probe> body)
        {
            using MemoryStream ms = new();
            Probe probe = format switch
            {
                "xlsx" => Probe.Of<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(XlsxWorkbookWriter.Create(ms, leaveOpen: true)),
                "xlsb" => Probe.Of<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(XlsbWorkbookWriter.Create(ms, leaveOpen: true)),
                _ => Probe.Of<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(XlsWorkbookWriter.Create(ms, leaveOpen: true)),
            };
            try
            {
                body(probe);
            }
            finally
            {
                probe.Abandon();
            }
        }

        private static byte[] Entry(byte[] package, string name)
        {
            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            using Stream data = zip.GetEntry(name)!.Open();
            using var copy = new MemoryStream();
            data.CopyTo(copy);
            return copy.ToArray();
        }

        /// <summary>The few workbook calls these tests make, over any of the three writers.</summary>
        private sealed class Probe
        {
            private Action<string> _addSheet = static _ => { };
            private Action<int, double> _width = static (_, _) => { };
            private Action<int> _style = static _ => { };
            private Func<CellStyle, int> _addStyle = static _ => 0;
            private Action _abandon = static () => { };

            internal static Probe Of<TWorkbook, TSheet, TRow>(TWorkbook workbook)
                where TWorkbook : IWorkbookWriter<TSheet>, IDisposable
                where TSheet : ISheetWriter<TRow>
                where TRow : IRowWriter
            {
                return new Probe
                {
                    _addSheet = name => workbook.AddSheet(name).End(),
                    _width = (column, width) =>
                    {
                        TSheet sheet = workbook.AddSheet("S1");
                        sheet.SetColumnWidth(column, width);
                        sheet.End();
                    },
                    _style = column =>
                    {
                        TSheet sheet = workbook.AddSheet("S2");
                        sheet.SetColumnStyle(column, 1);
                        sheet.End();
                    },
                    _addStyle = workbook.AddStyle,
                    _abandon = () =>
                    {
                        try
                        {
                            workbook.Dispose();
                        }
                        catch (InvalidOperationException)
                        {
                            // A workbook left without a finished sheet refuses to end; the probe only checks arguments.
                        }
                    },
                };
            }

            internal void AddSheet(string name)
            {
                _addSheet(name);
            }

            internal void Width(int column, double width)
            {
                _width(column, width);
            }

            internal void Style(int column)
            {
                _style(column);
            }

            internal int AddStyle(CellStyle style)
            {
                return _addStyle(style);
            }

            internal void Abandon()
            {
                _abandon();
            }
        }
    }
}
