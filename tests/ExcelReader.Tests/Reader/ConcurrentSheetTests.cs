using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader
{
    public sealed class ConcurrentSheetTests : IDisposable
    {
        public enum Format
        {
            Xlsx,
            Xlsb,
            Xls,
        }

        public enum Source
        {
            Memory,
            Stream,
            Path,
        }

        private const int Sheets = 6;
        private const int Rows = 3000;

        private readonly List<string> _paths = [];

        public void Dispose()
        {
            foreach (string path in _paths)
            {
                File.Delete(path);
            }
        }

        public static TheoryData<Format, Source> Cases
        {
            get
            {
                TheoryData<Format, Source> data = [];
                foreach (Format format in Enum.GetValues<Format>())
                {
                    foreach (Source source in Enum.GetValues<Source>())
                    {
                        data.Add(format, source);
                    }
                }
                return data;
            }
        }

        private static void Fill<TSheet, TRow>(IWorkbookWriter<TSheet> workbook)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            for (int s = 0; s < Sheets; s++)
            {
                TSheet sheet = workbook.AddSheet("sheet" + s.ToString(CultureInfo.InvariantCulture));
                for (int r = 0; r < Rows; r++)
                {
                    using TRow row = sheet.StartRow();
                    row.Write("s" + s.ToString(CultureInfo.InvariantCulture) + "-name-" + (r % 200).ToString(CultureInfo.InvariantCulture));
                    row.Write((s * 1_000_000) + r);
                    row.Write("shared-" + (r % 17).ToString(CultureInfo.InvariantCulture));
                }
                sheet.End();
                sheet.Dispose();
            }
            workbook.End();
        }

        private static byte[] Build(Format format)
        {
            using MemoryStream buffer = new();
            switch (format)
            {
                case Format.Xlsx:
                    using (XlsxWorkbookWriter xlsx = XlsxWorkbookWriter.Create(buffer, leaveOpen: true, new XlsxWriterOptions { UseSharedStrings = true }))
                    {
                        Fill<XlsxSheetWriter, XlsxRowWriter>(xlsx);
                    }
                    break;
                case Format.Xlsb:
                    using (XlsbWorkbookWriter xlsb = XlsbWorkbookWriter.Create(buffer, leaveOpen: true, new XlsbWriterOptions { UseSharedStrings = true }))
                    {
                        Fill<XlsbSheetWriter, XlsbRowWriter>(xlsb);
                    }
                    break;
                default:
                    using (XlsWorkbookWriter xls = XlsWorkbookWriter.Create(buffer, leaveOpen: true))
                    {
                        Fill<XlsSheetWriter, XlsRowWriter>(xls);
                    }
                    break;
            }
            return buffer.ToArray();
        }

        private IExcelWorkbook Open(Format format, Source source, byte[] bytes)
        {
            if (source == Source.Memory)
            {
                return Excel.Open(bytes);
            }
            if (source == Source.Stream)
            {
                return Excel.Open(new MemoryStream(bytes, writable: false), leaveOpen: false);
            }
            string path = Path.Combine(Path.GetTempPath(), "excelreader-concurrent-" + Guid.NewGuid().ToString("N") + "." + format.ToString());
            File.WriteAllBytes(path, bytes);
            _paths.Add(path);
            return Excel.Open(path);
        }

        private static List<string> Drain(IExcelRowEnumerator e)
        {
            List<string> lines = [];
            while (e.MoveNext())
            {
                Row row = e.Current;
                lines.Add(row[0].GetString() + "|" + row[1].GetString() + "|" + row[2].GetString());
            }
            return lines;
        }

        private static List<string>[] ReadSequentially(IExcelWorkbook reader)
        {
            List<string>[] sheets = new List<string>[reader.SheetCount];
            for (int s = 0; s < sheets.Length; s++)
            {
                using IExcelRowEnumerator e = reader.SheetAt(s).GetEnumerator();
                sheets[s] = Drain(e);
            }
            return sheets;
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void Sheets_Read_In_Parallel_Match_A_Sequential_Read(Format format, Source source)
        {
            byte[] bytes = Build(format);
            List<string>[] expected;
            using (IExcelWorkbook reference = Excel.Open(bytes))
            {
                expected = ReadSequentially(reference);
            }
            Assert.Equal(Sheets, expected.Length);
            Assert.All(expected, sheet => Assert.Equal(Rows, sheet.Count));

            using IExcelWorkbook reader = Open(format, source, bytes);
            IExcelRowEnumerator[] enumerators = new IExcelRowEnumerator[Sheets];
            for (int s = 0; s < Sheets; s++)
            {
                enumerators[s] = reader.SheetAt(s).GetEnumerator();
            }

            List<string>[] actual = new List<string>[Sheets];
            Parallel.For(0, Sheets, new ParallelOptions { MaxDegreeOfParallelism = Sheets }, s =>
            {
                using IExcelRowEnumerator e = enumerators[s];
                actual[s] = Drain(e);
            });

            for (int s = 0; s < Sheets; s++)
            {
                Assert.Equal(expected[s], actual[s]);
            }
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public async Task Sheets_Read_In_Parallel_Asynchronously_Match_A_Sequential_Read(Format format, Source source)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = Build(format);
            List<string>[] expected;
            using (IExcelWorkbook reference = Excel.Open(bytes))
            {
                expected = ReadSequentially(reference);
            }

            using IExcelWorkbook reader = Open(format, source, bytes);
            IExcelRowEnumerator[] enumerators = new IExcelRowEnumerator[Sheets];
            for (int s = 0; s < Sheets; s++)
            {
                enumerators[s] = reader.SheetAt(s).GetAsyncEnumerator(ct);
            }

            Task<List<string>>[] tasks = new Task<List<string>>[Sheets];
            for (int s = 0; s < Sheets; s++)
            {
                IExcelRowEnumerator e = enumerators[s];
                tasks[s] = Task.Run(
                    async () =>
                    {
                        List<string> lines = [];
                        await using (e)
                        {
                            while (await e.MoveNextAsync())
                            {
                                Row row = e.Current;
                                lines.Add(row[0].GetString() + "|" + row[1].GetString() + "|" + row[2].GetString());
                            }
                        }
                        return lines;
                    },
                    ct);
            }
            List<string>[] actual = await Task.WhenAll(tasks);

            for (int s = 0; s < Sheets; s++)
            {
                Assert.Equal(expected[s], actual[s]);
            }
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void Threads_Reading_One_Sheet_Get_The_Same_String_Instances(Format format, Source source)
        {
            byte[] bytes = Build(format);
            using IExcelWorkbook reader = Open(format, source, bytes);
            IExcelSheet sheet = reader.FirstSheet;

            const int Threads = 8;
            IExcelRowEnumerator[] enumerators = new IExcelRowEnumerator[Threads];
            for (int t = 0; t < Threads; t++)
            {
                enumerators[t] = sheet.GetEnumerator();
            }

            string[][] thirdColumn = new string[Threads][];
            Parallel.For(0, Threads, new ParallelOptions { MaxDegreeOfParallelism = Threads }, t =>
            {
                using IExcelRowEnumerator e = enumerators[t];
                string[] values = new string[Rows];
                int r = 0;
                while (e.MoveNext())
                {
                    values[r++] = e.Current[2].GetString();
                }
                thirdColumn[t] = values;
            });

            for (int r = 0; r < Rows; r++)
            {
                Assert.Equal("shared-" + (r % 17).ToString(CultureInfo.InvariantCulture), thirdColumn[0][r]);
                for (int t = 1; t < Threads; t++)
                {
                    if (format == Format.Xls)
                    {
                        Assert.Equal(thirdColumn[0][r], thirdColumn[t][r]);
                    }
                    else
                    {
                        Assert.Same(thirdColumn[0][r], thirdColumn[t][r]);
                    }
                }
            }
        }
    }
}
