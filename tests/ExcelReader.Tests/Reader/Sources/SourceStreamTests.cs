using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Sources;
using ExcelReader.Tests.Crypto;

namespace ExcelReader.Tests.Reader.Sources
{
    public sealed class SourceStreamTests
    {
        public static TheoryData<ConcurrentSheetTests.Format, ExcelFileFormat> Formats => new()
        {
            { ConcurrentSheetTests.Format.Xlsx, ExcelFileFormat.Xlsx },
            { ConcurrentSheetTests.Format.Xlsb, ExcelFileFormat.Xlsb },
            { ConcurrentSheetTests.Format.Xls, ExcelFileFormat.Xls },
            { ConcurrentSheetTests.Format.Xlsx, ExcelFileFormat.Unknown },
            { ConcurrentSheetTests.Format.Xls, ExcelFileFormat.Unknown },
        };

        public static TheoryData<string> EncryptedNames => new(EncryptedFixtures.All);

        private static List<string> Drain(IExcelSheet sheet)
        {
            List<string> lines = [];
            using IExcelRowEnumerator e = sheet.GetEnumerator();
            while (e.MoveNext())
            {
                StringBuilder line = new();
                foreach (RowCell cell in e.Current.Cells)
                {
                    line.Append(cell.Value.GetString()).Append('|');
                }
                lines.Add(line.ToString());
            }
            return lines;
        }

        [Fact]
        public void FromStream_Unwraps_A_SourceStream_That_Owns_Its_Source()
        {
            CountingSource source = new(new byte[100]);
            using ByteSource unwrapped = ByteSource.FromStream(new SourceStream(source), leaveOpen: false);
            Assert.Same(source, unwrapped);
        }

        [Fact]
        public void FromStream_Does_Not_Unwrap_A_Stream_The_Caller_Keeps()
        {
            using SourceStream stream = new(new CountingSource(new byte[100]));
            using ByteSource wrapped = ByteSource.FromStream(stream, leaveOpen: true);
            Assert.IsType<StreamByteSource>(wrapped);
        }

        [Fact]
        public void A_Taken_Source_Belongs_To_Its_New_Owner()
        {
            CountingSource source = new(new byte[100]);
            SourceStream stream = new(source);
            ByteSource unwrapped = ByteSource.FromStream(stream, leaveOpen: false);
            stream.Dispose();
            Assert.Equal(0, source.Disposals);
            unwrapped.Dispose();
            Assert.Equal(1, source.Disposals);
        }

        [Fact]
        public void Disposing_A_Stream_Whose_Source_Was_Not_Taken_Disposes_The_Source()
        {
            CountingSource source = new(new byte[100]);
            new SourceStream(source).Dispose();
            Assert.Equal(1, source.Disposals);
        }

        [Fact]
        public void Seek_And_Read_Follow_The_Source()
        {
            byte[] bytes = [.. Enumerable.Range(0, 256).Select(i => (byte)i)];
            using SourceStream stream = new(new CountingSource(bytes));
            Assert.True(stream.CanSeek);
            Assert.Equal(256, stream.Length);
            stream.Seek(-6, SeekOrigin.End);
            byte[] tail = new byte[10];
            Assert.Equal(6, stream.Read(tail));
            Assert.Equal(bytes[250..], tail[..6]);
            Assert.Equal(256, stream.Position);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public void Sheets_Opened_Over_A_SourceStream_Read_In_Parallel_What_Memory_Reads(
            ConcurrentSheetTests.Format format, ExcelFileFormat requested)
        {
            byte[] bytes = ConcurrentSheetTests.Build(format);
            List<string>[] expected;
            using (IExcelWorkbook reference = Excel.Open(bytes))
            {
                expected = [.. Enumerable.Range(0, reference.SheetCount).Select(s => Drain(reference.SheetAt(s)))];
            }

            CountingSource source = new(bytes);
            using (IExcelWorkbook workbook = Excel.Open(new SourceStream(source), requested, leaveOpen: false))
            {
                List<string>[] parallel = new List<string>[expected.Length * 2];
                Parallel.For(0, parallel.Length, i => parallel[i] = Drain(workbook.SheetAt(i % expected.Length)));
                for (int i = 0; i < parallel.Length; i++)
                {
                    Assert.Equal(expected[i % expected.Length], parallel[i]);
                }
            }
            Assert.Equal(1, source.Disposals);
        }

        [Fact]
        public void A_Csv_Over_A_SourceStream_Reads_More_Than_Once()
        {
            byte[] bytes = Encoding.UTF8.GetBytes("a,b,c\n1,2,3\n4,5,6\n");
            using IExcelWorkbook csv = Excel.Open(new SourceStream(new CountingSource(bytes)), ExcelFileFormat.Csv, leaveOpen: false);
            Assert.Equal(["a|b|c|", "1|2|3|", "4|5|6|"], Drain(csv.FirstSheet));
            Assert.Equal(Drain(csv.FirstSheet), Drain(csv.FirstSheet));
        }

        [Theory]
        [MemberData(nameof(EncryptedNames))]
        public void An_Encrypted_Package_Over_A_SourceStream_Opens_With_Its_Password(string name)
        {
            ExcelReaderOptions options = new() { Password = EncryptedFixtures.PasswordFor(name) };
            List<string> expected;
            using (IExcelWorkbook plain = Excel.Open(EncryptedFixtures.PlainBytes(name)))
            {
                expected = Drain(plain.FirstSheet);
            }
            using IExcelWorkbook workbook = Excel.Open(
                new SourceStream(new CountingSource(EncryptedFixtures.Bytes(name))), ExcelFileFormat.Unknown, leaveOpen: false, options);
            Assert.Equal(expected, Drain(workbook.FirstSheet));
        }
    }
}
