using System.Buffers.Binary;
using System.IO.Compression;
using System.Xml.Linq;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Hand-authored XML/ZIP fragments mimicking known producer quirks — not files actually exported by
    // those producers. See RealWorldXlsxCorpusTests for tests against genuine producer-exported binaries.
    public class XlsxProducerDialectShapeTests
    {
        private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PackageRelationshipsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string WorkbookRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string WorksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        private const string StylesRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        private const string SharedStringsRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

        public static IEnumerable<object[]> ProducerFixtures
        {
            get
            {
                yield return
                [
                    new ProducerFixture(
                        "ClosedXML-like dimensions, ignored views, and empty typed cells",
                        """
                        <dimension ref="A1:D2"/>
                        <sheetViews><sheetView workbookViewId="0"/></sheetViews>
                        <sheetFormatPr defaultRowHeight="15"/>
                        <sheetData>
                            <row r="1" spans="1:4" x14ac:dyDescent="0.25">
                                <c r="A1" t="s"><v>0</v></c>
                                <c r="B1"/>
                                <c r="C1" t="n"><v>123.45</v></c>
                                <c r="D1" t="b"><v>1</v></c>
                            </row>
                        </sheetData>
                        """,
                        "<si><t>ClosedXML</t></si>",
                        [
                            new ExpectedCell(0, 0, CellType.ExcelString, "ClosedXML"),
                            new ExpectedCell(0, 1, CellType.Empty, ""),
                            new ExpectedCell(0, 2, CellType.Number, "123.45"),
                            new ExpectedCell(0, 3, CellType.Boolean, "1"),
                        ])
                ];

                yield return
                [
                    new ProducerFixture(
                        "SheetJS-like dense refs, styled numbers, formula caches, and gaps",
                        """
                        <sheetData>
                            <row r="5">
                                <c r="A5" s="1"><v>45292</v></c>
                                <c r="C5" t="str"><f>CONCAT("A","B")</f><v>AB</v></c>
                                <c r="E5" t="e"><v>#N/A</v></c>
                            </row>
                        </sheetData>
                        <phoneticPr fontId="1" type="noConversion"/>
                        """,
                        null,
                        """<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>""",
                        [
                            new ExpectedCell(0, 0, CellType.Date, "45292"),
                            new ExpectedCell(0, 1, CellType.Empty, ""),
                            new ExpectedCell(0, 2, CellType.Formula, "AB"),
                            new ExpectedCell(0, 3, CellType.Empty, ""),
                            new ExpectedCell(0, 4, CellType.Error, "#N/A"),
                        ])
                ];

                yield return
                [
                    new ProducerFixture(
                        "Numbers-like preserved whitespace, rich shared strings, and extension lists",
                        """
                        <sheetData>
                            <row r="1">
                                <c r="A1" t="inlineStr"><is><t xml:space="preserve">  padded  </t></is></c>
                                <c r="B1" t="s"><v>0</v></c>
                            </row>
                        </sheetData>
                        <extLst><ext uri="{interop-fixture}"><ignored value="true"/></ext></extLst>
                        """,
                        "<si><r><rPr><b/></rPr><t>rich</t></r><r><t xml:space=\"preserve\"> text</t></r></si>",
                        [
                            new ExpectedCell(0, 0, CellType.ExcelString, "  padded  "),
                            new ExpectedCell(0, 1, CellType.ExcelString, "rich text"),
                        ])
                ];

                yield return
                [
                    new ProducerFixture(
                        "Google-Sheets-like ISO-8601 date cells (t=\"d\") alongside typed strings",
                        """
                        <sheetData>
                            <row r="1">
                                <c r="A1" t="d"><v>2024-01-01</v></c>
                                <c r="B1" t="d"><v>2024-01-01T00:00:00</v></c>
                                <c r="C1" t="s"><v>0</v></c>
                            </row>
                        </sheetData>
                        """,
                        "<si><t>label</t></si>",
                        [
                            new ExpectedCell(0, 0, CellType.Date, "45292"),
                            new ExpectedCell(0, 1, CellType.Date, "45292"),
                            new ExpectedCell(0, 2, CellType.ExcelString, "label"),
                        ])
                ];

                yield return
                [
                    new ProducerFixture(
                        "Aspose/Java-like namespace-prefixed worksheet with an unprefixed workbook part",
                        """
                        <x:sheetData>
                            <x:row r="1">
                                <x:c r="A1"><x:v>3.14</x:v></x:c>
                                <x:c r="B1" t="s"><x:v>0</x:v></x:c>
                                <x:c r="C1" t="inlineStr"><x:is><x:t>inline</x:t></x:is></x:c>
                            </x:row>
                        </x:sheetData>
                        """,
                        "<x:si><x:t>shared</x:t></x:si>",
                        [
                            new ExpectedCell(0, 0, CellType.Number, "3.14"),
                            new ExpectedCell(0, 1, CellType.ExcelString, "shared"),
                            new ExpectedCell(0, 2, CellType.ExcelString, "inline"),
                        ])
                    { Prefix = "x" }
                ];
            }
        }

        [Theory]
        [MemberData(nameof(ProducerFixtures))]
        public async Task ReaderHandlesProducerShapedXlsxFixtures(ProducerFixture fixture)
        {
            await using MemoryStream workbook = BuildProducerFixture(fixture);
            await using XlsxReader reader = await Excel.FromAsync(
                workbook,
                ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator rows = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.True(rows.MoveNext());
            foreach (ExpectedCell expected in fixture.Expected)
            {
                Assert.Equal(expected.Type, rows.Current[expected.Column].Type);
                Assert.Equal(expected.Value, rows.Current[expected.Column].GetString());
            }
            Assert.False(rows.MoveNext());
        }

        [Fact]
        public async Task XlsxWriterEmitsInteroperableOpenXmlPackageShape()
        {
            await using MemoryStream workbook = await WriteXlsxAsync(async wb =>
            {
                XlsxSheetWriter first = wb.AddSheet("Sales & Ops");
                await first.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsxRowWriter row = await first.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("<North>");
                    row.Write(42);
                    row.Write(true);
                }
                await first.EndAsync(TestContext.Current.CancellationToken);

                XlsxSheetWriter second = wb.AddSheet("Dates");
                await second.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsxRowWriter row = await second.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write(new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified));
                }
                await second.EndAsync(TestContext.Current.CancellationToken);
            });

            using var zip = new ZipArchive(workbook, ZipArchiveMode.Read, leaveOpen: true);
            AssertZipEntries(zip,
                "[Content_Types].xml",
                "_rels/.rels",
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "xl/styles.xml",
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml");

            XDocument rootRels = ReadXml(zip, "_rels/.rels");
            XElement rootRel = AssertSingleRelationship(rootRels, "rId1");
            Assert.Equal(WorkbookRelType, rootRel.Attribute("Type")?.Value);
            Assert.Equal("xl/workbook.xml", rootRel.Attribute("Target")?.Value);

            XDocument workbookXml = ReadXml(zip, "xl/workbook.xml");
            XNamespace s = SpreadsheetNs;
            XNamespace r = RelationshipsNs;
            XElement[] sheets = [.. workbookXml.Root!.Element(s + "sheets")!.Elements(s + "sheet")];
            Assert.Collection(
                sheets,
                sheet =>
                {
                    Assert.Equal("Sales & Ops", sheet.Attribute("name")?.Value);
                    Assert.Equal("rId2", sheet.Attribute(r + "id")?.Value);
                },
                sheet =>
                {
                    Assert.Equal("Dates", sheet.Attribute("name")?.Value);
                    Assert.Equal("rId3", sheet.Attribute(r + "id")?.Value);
                });

            XDocument workbookRels = ReadXml(zip, "xl/_rels/workbook.xml.rels");
            AssertRelationship(workbookRels, "rId1", StylesRelType, "styles.xml");
            AssertRelationship(workbookRels, "rId2", WorksheetRelType, "worksheets/sheet1.xml");
            AssertRelationship(workbookRels, "rId3", WorksheetRelType, "worksheets/sheet2.xml");

            XDocument contentTypes = ReadXml(zip, "[Content_Types].xml");
            AssertOverride(contentTypes, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            AssertOverride(contentTypes, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            AssertOverride(contentTypes, "/xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            AssertOverride(contentTypes, "/xl/worksheets/sheet2.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");

            XDocument sheet1 = ReadXml(zip, "xl/worksheets/sheet1.xml");
            XElement[] cells = [.. sheet1.Root!.Element(s + "sheetData")!.Element(s + "row")!.Elements(s + "c")];
            Assert.Equal("inlineStr", cells[0].Attribute("t")?.Value);
            Assert.Equal("<North>", cells[0].Element(s + "is")!.Element(s + "t")!.Value);
            Assert.Equal("42", cells[1].Element(s + "v")!.Value);
            Assert.Equal("b", cells[2].Attribute("t")?.Value);
            Assert.Equal("1", cells[2].Element(s + "v")!.Value);
        }

        [Fact]
        public async Task XlsbWriterEmitsInteroperableBinaryPackageShape()
        {
            await using MemoryStream workbook = await WriteXlsbAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("BinaryData");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("value");
                    row.Write(42);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            using var zip = new ZipArchive(workbook, ZipArchiveMode.Read, leaveOpen: true);
            AssertZipEntries(zip,
                "[Content_Types].xml",
                "_rels/.rels",
                "docProps/app.xml",
                "xl/workbook.bin",
                "xl/_rels/workbook.bin.rels",
                "xl/styles.bin",
                "xl/sharedStrings.bin",
                "xl/worksheets/sheet1.bin");

            Assert.True(zip.GetEntry("xl/workbook.bin")!.Length > 0);
            Assert.True(zip.GetEntry("xl/styles.bin")!.Length > 0);
            Assert.True(zip.GetEntry("xl/worksheets/sheet1.bin")!.Length > 0);

            XDocument rootRels = ReadXml(zip, "_rels/.rels");
            XElement rootRel = AssertSingleRelationship(rootRels, "wb");
            Assert.Equal(WorkbookRelType, rootRel.Attribute("Type")?.Value);
            Assert.Equal("xl/workbook.bin", rootRel.Attribute("Target")?.Value);

            XDocument workbookRels = ReadXml(zip, "xl/_rels/workbook.bin.rels");
            AssertRelationship(workbookRels, "s", StylesRelType, "styles.bin");
            AssertRelationship(workbookRels, "ss", SharedStringsRelType, "sharedStrings.bin");
            AssertRelationship(workbookRels, "s1", WorksheetRelType, "worksheets/sheet1.bin");

            XDocument contentTypes = ReadXml(zip, "[Content_Types].xml");
            AssertDefault(contentTypes, "rels", "application/vnd.openxmlformats-package.relationships+xml");
            AssertDefault(contentTypes, "bin", "application/vnd.ms-excel.sheet.binary.macroEnabled.main");
            AssertOverride(contentTypes, "/xl/workbook.bin", "application/vnd.ms-excel.sheet.binary.macroEnabled.main");
            AssertOverride(contentTypes, "/xl/styles.bin", "application/vnd.ms-excel.styles");
            AssertOverride(contentTypes, "/xl/sharedStrings.bin", "application/vnd.ms-excel.sharedStrings");
            AssertOverride(contentTypes, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            AssertOverride(contentTypes, "/xl/worksheets/sheet1.bin", "application/vnd.ms-excel.worksheet");
        }

        [Fact]
        public async Task XlsWriterEmitsOleCompoundWorkbookStream()
        {
            byte[] bytes = await WriteXlsAsync(wb =>
            {
                XlsSheetWriter sheet = wb.AddSheet("Legacy");
                sheet.Start();
                using (XlsRowWriter row = sheet.StartRow())
                {
                    row.Write("value");
                    row.Write(42);
                }
                sheet.End();
            });

            byte[] oleSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
            byte[] takenOleSignature = [.. bytes.Take(oleSignature.Length)];
            Assert.Equal(oleSignature, takenOleSignature);
            Assert.Equal(0x003E, BitConverter.ToUInt16(bytes, 0x18));
            Assert.Equal(0x0003, BitConverter.ToUInt16(bytes, 0x1A));

            int sectorSize = 1 << BitConverter.ToUInt16(bytes, 0x1E);
            int firstDirectorySector = BitConverter.ToInt32(bytes, 0x30);
            int directoryOffset = 512 + (firstDirectorySector * sectorSize);
            Assert.Contains("Root Entry", ReadDirectoryNames(bytes, directoryOffset, sectorSize));
            Assert.Contains("Workbook", ReadDirectoryNames(bytes, directoryOffset, sectorSize));
            Assert.Equal(1, BitConverter.ToInt32(bytes, directoryOffset + 76));
            Assert.Equal(unchecked((int)0xFFFFFFFF), BitConverter.ToInt32(bytes, directoryOffset + 128 + 68));
            Assert.Equal(unchecked((int)0xFFFFFFFF), BitConverter.ToInt32(bytes, directoryOffset + 128 + 72));
            Assert.Equal(unchecked((int)0xFFFFFFFF), BitConverter.ToInt32(bytes, directoryOffset + 128 + 76));

            using var reader = Excel.FromXls(new MemoryStream(bytes));
            Assert.Equal("Legacy", reader.SheetName);
        }

        private static MemoryStream BuildProducerFixture(ProducerFixture fixture)
        {
            // A fixture may prefix its worksheet/shared-strings elements (e.g. <x:row>) while the workbook
            // part stays unprefixed — the mixed shape where a prefixed worksheet previously read as zero rows.
            string pfx = fixture.Prefix is null ? "" : fixture.Prefix + ":";
            string wsNs = fixture.Prefix is null
                ? $"""xmlns="{SpreadsheetNs}" """
                : $"""xmlns:{fixture.Prefix}="{SpreadsheetNs}" """;
            string worksheet =
                $"""<{pfx}worksheet {wsNs}xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">{fixture.WorksheetInnerXml}</{pfx}worksheet>""";
            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteText(zip, "xl/worksheets/sheet1.xml", worksheet);
                WriteText(zip, "xl/workbook.xml",
                    $"""<workbook xmlns="{SpreadsheetNs}" xmlns:r="{RelationshipsNs}"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""");
                WriteText(zip, "xl/_rels/workbook.xml.rels",
                    $"""<Relationships xmlns="{PackageRelationshipsNs}"><Relationship Id="rId1" Type="{WorksheetRelType}" Target="worksheets/sheet1.xml"/></Relationships>""");
                if (fixture.SharedStringsXml is not null)
                {
                    WriteText(zip, "xl/sharedStrings.xml", $"""<{pfx}sst {wsNs.TrimEnd()}>{fixture.SharedStringsXml}</{pfx}sst>""");
                }
                if (fixture.StylesXml is not null)
                {
                    string styles = fixture.StylesXml.Replace("<styleSheet>", $"""<styleSheet xmlns="{SpreadsheetNs}">""", StringComparison.Ordinal);
                    WriteText(zip, "xl/styles.xml", styles);
                }
            }
            ms.Position = 0;
            return ms;
        }

        private static async Task<MemoryStream> WriteXlsxAsync(Func<XlsxWorkbookWriter, Task> build)
        {
            var ms = new MemoryStream();
            await using (XlsxWorkbookWriter writer = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                await writer.StartAsync(TestContext.Current.CancellationToken);
                await build(writer);
                await writer.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            return ms;
        }

        private static async Task<MemoryStream> WriteXlsbAsync(Func<XlsbWorkbookWriter, Task> build)
        {
            var ms = new MemoryStream();
            await using (XlsbWorkbookWriter writer = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                await writer.StartAsync(TestContext.Current.CancellationToken);
                await build(writer);
                await writer.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            return ms;
        }

        private static async Task<byte[]> WriteXlsAsync(Action<XlsWorkbookWriter> build)
        {
            var ms = new MemoryStream();
            await using (XlsWorkbookWriter writer = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                writer.Start();
                build(writer);
                await writer.EndAsync(TestContext.Current.CancellationToken);
            }
            return ms.ToArray();
        }

        private static void WriteText(ZipArchive zip, string name, string content)
        {
            using Stream stream = zip.CreateEntry(name).Open();
            using var writer = new StreamWriter(stream);
            writer.Write(content);
        }

        private static XDocument ReadXml(ZipArchive zip, string entryName)
        {
            using Stream stream = zip.GetEntry(entryName)!.Open();
            return XDocument.Load(stream);
        }

        private static void AssertZipEntries(ZipArchive zip, params string[] names)
        {
            foreach (string name in names)
            {
                Assert.NotNull(zip.GetEntry(name));
            }
        }

        private static XElement AssertSingleRelationship(XDocument document, string id)
        {
            XNamespace rel = PackageRelationshipsNs;
            return Assert.Single(
                document.Root!.Elements(rel + "Relationship"),
                e => string.Equals(e.Attribute("Id")?.Value, id, StringComparison.Ordinal));
        }

        private static void AssertRelationship(XDocument document, string id, string type, string target)
        {
            XElement rel = AssertSingleRelationship(document, id);
            Assert.Equal(type, rel.Attribute("Type")?.Value);
            Assert.Equal(target, rel.Attribute("Target")?.Value);
        }

        private static void AssertDefault(XDocument document, string extension, string contentType)
        {
            XNamespace ct = ContentTypesNs;
            XElement match = Assert.Single(
                document.Root!.Elements(ct + "Default"),
                e => string.Equals(e.Attribute("Extension")?.Value, extension, StringComparison.Ordinal));
            Assert.Equal(contentType, match.Attribute("ContentType")?.Value);
        }

        private static void AssertOverride(XDocument document, string partName, string contentType)
        {
            XNamespace ct = ContentTypesNs;
            XElement match = Assert.Single(
                document.Root!.Elements(ct + "Override"),
                e => string.Equals(e.Attribute("PartName")?.Value, partName, StringComparison.Ordinal));
            Assert.Equal(contentType, match.Attribute("ContentType")?.Value);
        }

        private static string[] ReadDirectoryNames(byte[] bytes, int directoryOffset, int length)
        {
            var names = new List<string>();
            for (int offset = directoryOffset; offset < directoryOffset + length; offset += 128)
            {
                ushort nameBytes = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 64, 2));
                if (nameBytes >= 2)
                {
                    names.Add(System.Text.Encoding.Unicode.GetString(bytes, offset, nameBytes - 2));
                }
            }
            return [.. names];
        }

        public sealed record ProducerFixture(
            string Name,
            string WorksheetInnerXml,
            string? SharedStringsXml,
            ExpectedCell[] Expected)
        {
            public ProducerFixture(
                string name,
                string worksheetInnerXml,
                string? sharedStringsXml,
                string? stylesXml,
                ExpectedCell[] expected)
                : this(name, worksheetInnerXml, sharedStringsXml, expected)
            {
                StylesXml = stylesXml;
            }

            public string? StylesXml { get; init; }

            // When set, the worksheet (and shared-strings) elements are namespace-prefixed, e.g. <x:row>.
            public string? Prefix { get; init; }

            public override string ToString()
            {
                return Name;
            }
        }

        public readonly record struct ExpectedCell(int Row, int Column, CellType Type, string Value);
    }
}
