using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Benchmarks
{
    internal static class Program
    {
        public static void Main()
        {
            BenchmarkRunner.Run<ReadBenchmark>();
        }

    }
    [MemoryDiagnoser]
    public class ReadBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private byte[] _workbook = [];

        [GlobalSetup]
        public void Setup()
        {
            _workbook = WorkbookGenerator.Build(Rows);
        }

        // Full read pass: iterate every row/cell and touch each value the way a consumer would.

        [Benchmark]
        public long ReadAllCells()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = Excel.From(ms);

            long acc = 0;
            foreach (var row in reader)
            {
                for (int c = 0; c < row.ColumnCount; c++)
                {
                    var cell = row[c];
                    switch (cell.Type)
                    {
                        case CellType.ExcelString:
                            acc += cell.Value.Length;
                            break;
                        case CellType.Number:
                            if (cell.TryParse<double>(null, out double n)) { acc += (long)n; }
                            break;
                        case CellType.Date:
                            if (cell.TryGetDateTime(out var d)) { acc += d.Ticks; }
                            break;
                        default:
                            break;
                    }
                }
            }
            return acc;
        }
    }

    // Generates a self-contained .xlsx in memory: a header row of shared strings plus `rows` data rows
    // of [shared string, integer, date, float] — enough variety to exercise all the reader's value paths.
    internal static class WorkbookGenerator
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        public static byte[] Build(int rows)
        {
            string[] pool = ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];

            var shared = new StringBuilder($"<sst xmlns=\"{Main}\" count=\"{pool.Length}\" uniqueCount=\"{pool.Length}\">");
            foreach (var s in pool)
            {
                shared.Append("<si><t>").Append(s).Append("</t></si>");
            }
            shared.Append("</sst>");

            var sheet = new StringBuilder($"<worksheet xmlns=\"{Main}\"><sheetData>");
            for (int r = 1; r <= rows; r++)
            {
                int si = r % pool.Length;
                double serial = 45292 + (r % 3650) + 0.25; // dates spread over ~10 years
                sheet.Append("<row r=\"").Append(r).Append("\">")
                    .Append("<c r=\"A").Append(r).Append("\" t=\"s\"><v>").Append(si).Append("</v></c>")
                    .Append("<c r=\"B").Append(r).Append("\"><v>").Append(r).Append("</v></c>")
                    .Append("<c r=\"C").Append(r).Append("\" s=\"1\"><v>").Append(serial.ToString("0.####", CultureInfo.InvariantCulture)).Append("</v></c>")
                    .Append("<c r=\"D").Append(r).Append("\"><v>").Append((r * 1.5).ToString("0.####", CultureInfo.InvariantCulture)).Append("</v></c>")
                    .Append("</row>");
            }
            sheet.Append("</sheetData></worksheet>");

            // cellXfs[1] -> builtin numFmtId 14 (date), used by column C above.
            const string styles =
                "<styleSheet xmlns=\"" + Main + "\"><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>";

            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                Write(zip, "xl/workbook.xml",
                    $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                Write(zip, "xl/_rels/workbook.xml.rels",
                    $"<Relationships xmlns=\"{PkgRel}\"><Relationship Id=\"rId1\" Type=\"x\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                Write(zip, "xl/worksheets/sheet1.xml", sheet.ToString());
                Write(zip, "xl/sharedStrings.xml", shared.ToString());
                Write(zip, "xl/styles.xml", styles);
            }
            return ms.ToArray();
        }

        private static void Write(ZipArchive zip, string name, string content)
        {
            using var s = zip.CreateEntry(name).Open();
            var bytes = Encoding.UTF8.GetBytes(content);
            s.Write(bytes, 0, bytes.Length);
        }
    }
}

