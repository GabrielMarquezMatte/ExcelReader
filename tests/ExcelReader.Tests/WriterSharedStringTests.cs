using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class WriterSharedStringTests
    {
        private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string PackageRelationshipsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string SharedStringsRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

        [Fact]
        public async Task XlsxWriterUsesInlineStringsByDefault()
        {
            await using MemoryStream workbook = await WriteXlsxAsync(useSharedStrings: false, async wb =>
            {
                SheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (RowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("repeat");
                    row.Write("repeat");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            using var zip = new ZipArchive(workbook, ZipArchiveMode.Read, leaveOpen: true);
            Assert.Null(zip.GetEntry("xl/sharedStrings.xml"));

            XDocument sheetXml = ReadXml(zip, "xl/worksheets/sheet1.xml");
            XNamespace ns = SpreadsheetNs;
            XElement[] cells = sheetXml.Root!.Element(ns + "sheetData")!.Element(ns + "row")!.Elements(ns + "c").ToArray();
            Assert.All(cells, cell => Assert.Equal("inlineStr", cell.Attribute("t")?.Value));
        }

        [Fact]
        public async Task XlsxWriterOptInSharedStringsDeduplicatesAndRoundTrips()
        {
            await using MemoryStream workbook = await WriteXlsxAsync(useSharedStrings: true, async wb =>
            {
                SheetWriter first = wb.AddSheet("First");
                await first.StartAsync(TestContext.Current.CancellationToken);
                await using (RowWriter row = await first.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("repeat");
                    row.Write("repeat");
                    row.Write(" spaced ");
                }
                await first.EndAsync(TestContext.Current.CancellationToken);

                SheetWriter second = wb.AddSheet("Second");
                await second.StartAsync(TestContext.Current.CancellationToken);
                await using (RowWriter row = await second.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("repeat");
                }
                await second.EndAsync(TestContext.Current.CancellationToken);
            });

            using (var zip = new ZipArchive(workbook, ZipArchiveMode.Read, leaveOpen: true))
            {
                XDocument shared = ReadXml(zip, "xl/sharedStrings.xml");
                XNamespace ns = SpreadsheetNs;
                Assert.Equal("4", shared.Root!.Attribute("count")?.Value);
                Assert.Equal("2", shared.Root.Attribute("uniqueCount")?.Value);
                XElement[] items = shared.Root.Elements(ns + "si").ToArray();
                Assert.Equal("repeat", items[0].Element(ns + "t")!.Value);
                XElement spaced = items[1].Element(ns + "t")!;
                Assert.Equal(" spaced ", spaced.Value);
                Assert.Equal("preserve", spaced.Attribute(XNamespace.Xml + "space")?.Value);

                XDocument sheet = ReadXml(zip, "xl/worksheets/sheet1.xml");
                XElement[] cells = sheet.Root!.Element(ns + "sheetData")!.Element(ns + "row")!.Elements(ns + "c").ToArray();
                Assert.Collection(
                    cells,
                    cell => AssertSharedCell(cell, "0"),
                    cell => AssertSharedCell(cell, "0"),
                    cell => AssertSharedCell(cell, "1"));

                XDocument rels = ReadXml(zip, "xl/_rels/workbook.xml.rels");
                AssertRelationship(rels, "rIdShared", SharedStringsRelType, "sharedStrings.xml");
                XDocument contentTypes = ReadXml(zip, "[Content_Types].xml");
                AssertOverride(contentTypes, "/xl/sharedStrings.xml", SharedStringsContentType);
            }

            workbook.Position = 0;
            await using XlsxReader reader = Excel.From(workbook);
            using XlsxReader.Enumerator rows = reader.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal("repeat", rows.Current[0].GetString());
            Assert.Equal("repeat", rows.Current[1].GetString());
            Assert.Equal(" spaced ", rows.Current[2].GetString());
            Assert.False(rows.MoveNext());

            reader.MoveToSheet(1);
            using XlsxReader.Enumerator secondRows = reader.GetEnumerator();
            Assert.True(secondRows.MoveNext());
            Assert.Equal("repeat", secondRows.Current[0].GetString());
        }

        [Fact]
        public async Task XlsbWriterUsesInlineStringsByDefault()
        {
            await using MemoryStream workbook = await WriteXlsbAsync(useSharedStrings: false, async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("repeat");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            using var zip = new ZipArchive(workbook, ZipArchiveMode.Read, leaveOpen: true);
            Assert.Equal(0, zip.GetEntry("xl/sharedStrings.bin")!.Length);
            byte[] sheetBytes = ReadEntry(zip, "xl/worksheets/sheet1.bin");
            Assert.Contains(ReadRecordIds(sheetBytes), id => id == Brt.CellSt);
            Assert.DoesNotContain(ReadRecordIds(sheetBytes), id => id == Brt.CellIsst);
        }

        [Fact]
        public async Task XlsbWriterOptInSharedStringsDeduplicatesAndRoundTrips()
        {
            await using MemoryStream workbook = await WriteXlsbAsync(useSharedStrings: true, async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("repeat");
                    row.Write("repeat");
                    row.Write("Café");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            using (var zip = new ZipArchive(workbook, ZipArchiveMode.Read, leaveOpen: true))
            {
                byte[] sharedBytes = ReadEntry(zip, "xl/sharedStrings.bin");
                var (flat, offsets) = XlsbSharedStrings.Parse(sharedBytes);
                Assert.Equal(3, offsets.Length);
                Assert.Equal("repeat", SharedAt(flat, offsets, 0));
                Assert.Equal("Café", SharedAt(flat, offsets, 1));

                byte[] sheetBytes = ReadEntry(zip, "xl/worksheets/sheet1.bin");
                Assert.Equal([0u, 0u, 1u], ReadCellIsstIndexes(sheetBytes));
                Assert.DoesNotContain(ReadRecordIds(sheetBytes), id => id == Brt.CellSt);
            }

            workbook.Position = 0;
            await using XlsbReader reader = Excel.FromXlsb(workbook);
            using XlsbReader.Enumerator rows = reader.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal(CellType.ExcelString, rows.Current[0].Type);
            Assert.Equal("repeat", rows.Current[0].GetString());
            Assert.Equal("repeat", rows.Current[1].GetString());
            Assert.Equal("Café", rows.Current[2].GetString());
            Assert.False(rows.MoveNext());
        }

        private static async Task<MemoryStream> WriteXlsxAsync(bool useSharedStrings, Func<WorkbookWriter, Task> build)
        {
            var ms = new MemoryStream();
            await using (WorkbookWriter writer = await WorkbookWriter.CreateAsync(
                ms,
                leaveOpen: true,
                useSharedStrings: useSharedStrings,
                ct: TestContext.Current.CancellationToken))
            {
                await writer.StartAsync(TestContext.Current.CancellationToken);
                await build(writer);
                await writer.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            return ms;
        }

        private static async Task<MemoryStream> WriteXlsbAsync(bool useSharedStrings, Func<XlsbWorkbookWriter, Task> build)
        {
            var ms = new MemoryStream();
            await using (XlsbWorkbookWriter writer = await XlsbWorkbookWriter.CreateAsync(
                ms,
                leaveOpen: true,
                useSharedStrings: useSharedStrings,
                ct: TestContext.Current.CancellationToken))
            {
                await writer.StartAsync(TestContext.Current.CancellationToken);
                await build(writer);
                await writer.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            return ms;
        }

        private static XDocument ReadXml(ZipArchive zip, string entryName)
        {
            using Stream stream = zip.GetEntry(entryName)!.Open();
            return XDocument.Load(stream);
        }

        private static byte[] ReadEntry(ZipArchive zip, string entryName)
        {
            using Stream stream = zip.GetEntry(entryName)!.Open();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }

        private static void AssertSharedCell(XElement cell, string index)
        {
            XNamespace ns = SpreadsheetNs;
            Assert.Equal("s", cell.Attribute("t")?.Value);
            Assert.Equal(index, cell.Element(ns + "v")?.Value);
        }

        private static void AssertRelationship(XDocument document, string id, string type, string target)
        {
            XNamespace rel = PackageRelationshipsNs;
            XElement match = Assert.Single(
                document.Root!.Elements(rel + "Relationship"),
                e => string.Equals(e.Attribute("Id")?.Value, id, StringComparison.Ordinal));
            Assert.Equal(type, match.Attribute("Type")?.Value);
            Assert.Equal(target, match.Attribute("Target")?.Value);
        }

        private static void AssertOverride(XDocument document, string partName, string contentType)
        {
            XNamespace ct = ContentTypesNs;
            XElement match = Assert.Single(
                document.Root!.Elements(ct + "Override"),
                e => string.Equals(e.Attribute("PartName")?.Value, partName, StringComparison.Ordinal));
            Assert.Equal(contentType, match.Attribute("ContentType")?.Value);
        }

        private static int[] ReadRecordIds(ReadOnlySpan<byte> data)
        {
            var ids = new List<int>();
            var reader = new Biff12RecordReader(data);
            while (reader.TryReadRecord(out int id, out _))
            {
                ids.Add(id);
            }
            return [.. ids];
        }

        private static uint[] ReadCellIsstIndexes(ReadOnlySpan<byte> data)
        {
            var indexes = new List<uint>();
            var reader = new Biff12RecordReader(data);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                if (id == Brt.CellIsst && payload.Length >= 12)
                {
                    indexes.Add(Biff12.ReadU32(payload, 8));
                }
            }
            return [.. indexes];
        }

        private static string SharedAt(byte[] flat, int[] offsets, int index)
        {
            return Encoding.UTF8.GetString(flat.AsSpan(offsets[index], offsets[index + 1] - offsets[index]));
        }
    }
}
