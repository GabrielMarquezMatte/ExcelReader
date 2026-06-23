using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using MiniExcelLibs;
using Sylvan.Data.Excel;

namespace ExcelReader.Benchmarks
{
    internal static class Program
    {
        public static void Main(string[] args)
        {
            BenchmarkRunner.Run<ReadBenchmark>(args: args);
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

        [Benchmark(Baseline = true)]
        public long ExcelReader()
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

        [Benchmark]
        public async Task<long> ExcelReaderAsync()
        {
            await using var ms = new MemoryStream(_workbook, writable: false);
            await using var reader = await Excel.FromAsync(ms);
            await using var e = await reader.GetAsyncEnumeratorAsync();
            long acc = 0;
            while (await e.MoveNextAsync())
            {
                var row = e.Current;
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

        [Benchmark]
        public long MiniExcel()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            long acc = 0;
            foreach (var row in ms.Query(useHeaderRow: false, excelType: ExcelType.XLSX))
            {
                var r = (IDictionary<string, object?>)row;
                foreach (var val in r.Values)
                {
                    switch (val)
                    {
                        case string s: acc += s.Length; break;
                        case double d: acc += (long)d; break;
                        case DateTime dt: acc += dt.Ticks; break;
                    }
                }
            }
            return acc;
        }

        [Benchmark]
        public long Sylvan()
        {
            using var ms = new MemoryStream(_workbook, writable: false);
            using var reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            long acc = 0;
            do
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) { continue; }
                        switch (reader.GetExcelDataType(i))
                        {
                            case ExcelDataType.String:
                                acc += reader.GetString(i).Length;
                                break;
                            case ExcelDataType.Numeric:
                                acc += (long)reader.GetDouble(i);
                                break;
                            case ExcelDataType.DateTime:
                                acc += reader.GetDateTime(i).Ticks;
                                break;
                            case ExcelDataType.Boolean:
                            case ExcelDataType.Error:
                            case ExcelDataType.Null:
                                break;
                        }
                    }
                }
            }
            while (reader.NextResult());
            return acc;
        }
    }

    // Generates a self-contained .xlsx in memory: `rows` data rows of
    // [shared string, integer, date, float] — exercises all reader value paths.
    internal static class WorkbookGenerator
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string CT = "http://schemas.openxmlformats.org/package/2006/content-types";

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
                Write(zip, "[Content_Types].xml",
                    $"<Types xmlns=\"{CT}\">" +
                    $"<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    $"<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    $"<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    $"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    $"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
                    $"<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    $"</Types>");
                Write(zip, "_rels/.rels",
                    $"<Relationships xmlns=\"{PkgRel}\"><Relationship Id=\"rId0\" Type=\"{Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                Write(zip, "xl/workbook.xml",
                    $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                Write(zip, "xl/_rels/workbook.xml.rels",
                    $"<Relationships xmlns=\"{PkgRel}\">" +
                    $"<Relationship Id=\"rId1\" Type=\"{Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    $"<Relationship Id=\"rId2\" Type=\"{Rel}/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
                    $"<Relationship Id=\"rId3\" Type=\"{Rel}/styles\" Target=\"styles.xml\"/>" +
                    $"</Relationships>");
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
