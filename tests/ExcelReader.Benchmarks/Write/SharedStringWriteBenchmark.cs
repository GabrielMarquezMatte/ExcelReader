using System.Globalization;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Benchmarks
{
    // Every text distinct, so the shared-string part is as large as the sheet that references it.
    [MemoryDiagnoser]
    public class SharedStringWriteBenchmark
    {
        [Params(100_000)]
        public int Texts { get; set; }

        private string[] _texts = [];

        [GlobalSetup]
        public void Setup()
        {
            _texts = [.. Enumerable.Range(0, Texts).Select(static i => i.ToString("D8", CultureInfo.InvariantCulture) + new string('x', 120))];
        }

        [Benchmark]
        public long Xlsx()
        {
            using MemoryStream ms = new(32 * 1024 * 1024);
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsxWriterOptions { UseSharedStrings = true }))
            {
                using XlsxSheetWriter sheet = workbook.AddSheet("S");
                foreach (string text in _texts)
                {
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write(text);
                }
                sheet.End();
                workbook.End();
            }
            return ms.Length;
        }

        [Benchmark]
        public long Xlsb()
        {
            using MemoryStream ms = new(32 * 1024 * 1024);
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsbWriterOptions { UseSharedStrings = true }))
            {
                using XlsbSheetWriter sheet = workbook.AddSheet("S");
                foreach (string text in _texts)
                {
                    using XlsbRowWriter row = sheet.StartRow();
                    row.Write(text);
                }
                sheet.End();
                workbook.End();
            }
            return ms.Length;
        }
    }
}
