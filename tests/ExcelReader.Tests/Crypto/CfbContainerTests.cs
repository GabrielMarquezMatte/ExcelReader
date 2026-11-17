using System.Buffers.Binary;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Tests.Reader.Xls;

namespace ExcelReader.Tests.Crypto
{
    public class CfbContainerTests
    {
        private static CfbContainer Open(string fixture)
        {
            var fs = File.OpenRead(EncryptedFixtures.Path_(fixture));
            return CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
        }

        [Fact]
        public void Should_Throw_When_The_Only_Fat_Sector_Points_Past_The_End_Of_The_File()
        {
            using MemoryStream built = XlsWorkbookBuilder.Build(sheets: [("S1", [["Alice", 1, true]])]);
            byte[] bytes = built.ToArray();
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(0x4C), 1_048_576);

            Assert.Throws<InvalidDataException>(
                () => CfbContainer.Parse(new MemoryStream(bytes), ownsSource: true, ExcelReaderOptions.Default));
        }

        private static byte[] FormulaStringsXls()
        {
            return File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "data", "excel-formula-strings.xls"));
        }

        private static Span<byte> DirectoryEntryBytes(byte[] bytes, int index)
        {
            int sectorSize = 1 << BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(0x1E));
            int directorySector = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(0x30));
            return bytes.AsSpan(((directorySector + 1) * sectorSize) + (index * 128), 128);
        }

        private static void SetLinks(byte[] bytes, int index, int left, int right, int child)
        {
            Span<byte> entry = DirectoryEntryBytes(bytes, index);
            BinaryPrimitives.WriteInt32LittleEndian(entry[68..], left);
            BinaryPrimitives.WriteInt32LittleEndian(entry[72..], right);
            BinaryPrimitives.WriteInt32LittleEndian(entry[76..], child);
        }

        // FAT entry i points to i + 1, so the directory chain walks sectors far past the end of the file.
        private static byte[] BuildLongFatChain(int fatSectors)
        {
            const int SectorSize = 512;
            const int EntriesPerSector = SectorSize / 4;
            int difatSectors = fatSectors > 109 ? ((fatSectors - 109) + 126) / 127 : 0;
            byte[] bytes = new byte[SectorSize * (1 + fatSectors + difatSectors)];
            Span<byte> header = bytes.AsSpan(0, SectorSize);
            ReadOnlySpan<byte> signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
            signature.CopyTo(header);
            BinaryPrimitives.WriteUInt16LittleEndian(header[0x18..], 0x3E);
            BinaryPrimitives.WriteUInt16LittleEndian(header[0x1A..], 3);
            BinaryPrimitives.WriteUInt16LittleEndian(header[0x1C..], 0xFFFE);
            BinaryPrimitives.WriteUInt16LittleEndian(header[0x1E..], 9);
            BinaryPrimitives.WriteUInt16LittleEndian(header[0x20..], 6);
            BinaryPrimitives.WriteInt32LittleEndian(header[0x2C..], fatSectors);
            BinaryPrimitives.WriteInt32LittleEndian(header[0x30..], 0);
            BinaryPrimitives.WriteInt32LittleEndian(header[0x38..], 4096);
            BinaryPrimitives.WriteInt32LittleEndian(header[0x3C..], -2);
            BinaryPrimitives.WriteInt32LittleEndian(header[0x40..], 0);
            BinaryPrimitives.WriteInt32LittleEndian(header[0x44..], difatSectors > 0 ? fatSectors : -2);
            BinaryPrimitives.WriteInt32LittleEndian(header[0x48..], difatSectors);
            for (int i = 0; i < 109; i++)
            {
                BinaryPrimitives.WriteInt32LittleEndian(header[(0x4C + (i * 4))..], i < fatSectors ? i : -1);
            }

            int nextFat = 109;
            for (int d = 0; d < difatSectors; d++)
            {
                Span<byte> sector = bytes.AsSpan(SectorSize * (1 + fatSectors + d), SectorSize);
                for (int j = 0; j < 127; j++)
                {
                    BinaryPrimitives.WriteInt32LittleEndian(sector[(j * 4)..], nextFat < fatSectors ? nextFat++ : -1);
                }
                BinaryPrimitives.WriteInt32LittleEndian(sector[(127 * 4)..], d + 1 < difatSectors ? fatSectors + d + 1 : -2);
            }

            int totalEntries = fatSectors * EntriesPerSector;
            for (int k = 0; k < fatSectors; k++)
            {
                Span<byte> sector = bytes.AsSpan(SectorSize * (1 + k), SectorSize);
                for (int j = 0; j < EntriesPerSector; j++)
                {
                    int index = (k * EntriesPerSector) + j;
                    BinaryPrimitives.WriteInt32LittleEndian(sector[(j * 4)..], index + 1 < totalEntries ? index + 1 : -2);
                }
            }
            return bytes;
        }

        [Fact]
        public void Should_Throw_Without_Allocating_The_Chain_When_The_Directory_Chain_Runs_Past_The_File()
        {
            byte[] bytes = BuildLongFatChain(fatSectors: 2048);

            long before = GC.GetAllocatedBytesForCurrentThread();
            Exception? thrown = Record.Exception(() => Excel.FromXls(bytes));
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

            Assert.IsType<InvalidDataException>(thrown);
            Assert.True(allocated < 16L * bytes.Length, $"file {bytes.Length:N0} bytes, allocated {allocated:N0} bytes");
        }

        [Fact]
        public void Should_Reject_As_Corrupt_When_The_Directory_Chain_Exceeds_Int32_Bytes()
        {
            byte[] bytes = BuildLongFatChain(fatSectors: 33_000);

            Assert.Equal(ExcelFileFormat.Xls, Excel.DetectFileFormat(bytes));
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(bytes));
        }

        [Fact]
        public void Should_Ignore_Streams_Inside_Embedded_Storages_When_Looking_Up_By_Name()
        {
            byte[] bytes = FormulaStringsXls();
            Span<byte> summary = DirectoryEntryBytes(bytes, 2);
            summary[66] = 1;
            SetLinks(bytes, 2, left: 1, right: -1, child: 3);
            Span<byte> nested = DirectoryEntryBytes(bytes, 3);
            nested[..64].Clear();
            System.Text.Encoding.Unicode.GetBytes("EncryptedPackage").CopyTo(nested);
            BinaryPrimitives.WriteUInt16LittleEndian(nested[64..], (ushort)(("EncryptedPackage".Length + 1) * 2));

            using (CfbContainer cfb = CfbContainer.Parse(new MemoryStream(bytes), ownsSource: true, ExcelReaderOptions.Default))
            {
                Assert.False(cfb.ContainsStream("EncryptedPackage"));
                Assert.True(cfb.ContainsStream("Workbook"));
            }
            Assert.Equal(ExcelFileFormat.Xls, Excel.DetectFileFormat(bytes));
            using IExcelWorkbook workbook = Excel.Open(new MemoryStream(bytes));
            Assert.IsType<XlsWorkbook>(workbook);
        }

        [Fact]
        public void Should_Throw_When_Directory_Sibling_Tree_Has_A_Cycle()
        {
            byte[] bytes = FormulaStringsXls();
            SetLinks(bytes, 1, left: 2, right: -1, child: -1);

            using CfbContainer cfb = CfbContainer.Parse(new MemoryStream(bytes), ownsSource: true, ExcelReaderOptions.Default);
            Assert.Throws<InvalidDataException>(() => cfb.ContainsStream("NoSuchStream"));
        }

        [Theory]
        [InlineData(4)]
        [InlineData(1_000_000)]
        [InlineData(int.MinValue)]
        public void Should_Not_Find_Streams_When_Root_Child_Is_Out_Of_Range(int child)
        {
            byte[] bytes = FormulaStringsXls();
            SetLinks(bytes, 0, left: -1, right: -1, child: child);

            using CfbContainer cfb = CfbContainer.Parse(new MemoryStream(bytes), ownsSource: true, ExcelReaderOptions.Default);
            Assert.False(cfb.ContainsStream("Workbook"));
        }

        [Fact]
        public void Should_Ignore_Stream_Size_High_Dword_When_Version_3_Container()
        {
            byte[] bytes = FormulaStringsXls();
            Span<byte> workbookEntry = DirectoryEntryBytes(bytes, 1);
            long size = BinaryPrimitives.ReadInt64LittleEndian(workbookEntry[120..]);
            BinaryPrimitives.WriteUInt32LittleEndian(workbookEntry[124..], 0xDEADBEEF);

            using (CfbContainer cfb = CfbContainer.Parse(new MemoryStream(bytes), ownsSource: true, ExcelReaderOptions.Default))
            {
                Assert.Equal(size, cfb.StreamLength("Workbook"));
            }
            using XlsWorkbook workbook = Excel.FromXls(new MemoryStream(bytes));
            using XlsWorkbook.Enumerator e = workbook.FirstSheet.GetEnumerator();
            Assert.True(e.MoveNext());
        }

        [Fact]
        public void Should_Find_Both_Encryption_Streams_When_Encrypted_Container()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            Assert.True(cfb.ContainsStream("EncryptionInfo"));
            Assert.True(cfb.ContainsStream("EncryptedPackage"));
            Assert.False(cfb.ContainsStream("Workbook"));
        }

        [Fact]
        public void Should_Read_EncryptionInfo_When_Below_Mini_Cutoff()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            byte[] info = cfb.ReadStream("EncryptionInfo", maxBytes: 64 * 1024);
            Assert.InRange(info.Length, 64, 64 * 1024);
            Assert.Equal(4, info[0] | (info[1] << 8));
            Assert.Equal(4, info[2] | (info[3] << 8));
        }

        [Fact]
        public void Should_Throw_When_Stream_Exceeds_Max_Bytes()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            Assert.Throws<ExcelLimitExceededException>(
                () => cfb.ReadStream("EncryptionInfo", maxBytes: 8));
        }

        [Fact]
        public void Should_Throw_When_Stream_Absent()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            Assert.Throws<InvalidDataException>(() => cfb.ReadStream("NoSuchStream", 1024));
        }

        [Fact]
        public void Should_Expose_Seekable_View_When_Opening_EncryptedPackage()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            using Stream view = cfb.OpenStreamView("EncryptedPackage");
            Assert.True(view.CanRead);
            Assert.True(view.CanSeek);
            Assert.False(view.CanWrite);
            Assert.Equal(cfb.StreamLength("EncryptedPackage"), view.Length);

            byte[] prefix = new byte[8];
            view.ReadExactly(prefix);
            long declared = BitConverter.ToInt64(prefix);
            Assert.InRange(declared, 1, view.Length);
        }

        [Fact]
        public void Should_Match_Sequential_Read_When_Seeking()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            using Stream view = cfb.OpenStreamView("EncryptedPackage");
            byte[] whole = new byte[view.Length];
            view.ReadExactly(whole);

            view.Seek(-64, SeekOrigin.End);
            byte[] tail = new byte[64];
            view.ReadExactly(tail);
            Assert.Equal(whole[^64..], tail);
        }
    }
}
