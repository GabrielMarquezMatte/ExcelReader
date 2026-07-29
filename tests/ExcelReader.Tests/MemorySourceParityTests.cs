using System.Buffers;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class MemorySourceParityTests
    {
        [Fact]
        public void CsvMemorySourceMatchesStreamSource()
        {
            byte[] bytes = Encoding.UTF8.GetBytes("name,age\nAda,37\n");

            using CsvReader streamReader = Excel.FromCsv(new MemoryStream(bytes, writable: false));
            using CsvReader memoryReader = Excel.FromCsv(bytes.AsMemory());

            AssertRowsEqual(streamReader, memoryReader);
        }

        [Fact]
        public void CsvMemorySourceMatchesStreamSource_ForSlicedAndNonArrayBackedBuffers()
        {
            byte[] bytes = Encoding.UTF8.GetBytes("name,age\nAda,37\n");
            ReadOnlyMemory<byte> sliced = bytes.AsMemory(1, bytes.Length - 2);
            var manager = new NonArrayMemoryManager(sliced.ToArray());

            using CsvReader streamReader = Excel.FromCsv(new MemoryStream(sliced.ToArray(), writable: false));
            using CsvReader memoryReader = Excel.FromCsv(sliced);
            using CsvReader managerReader = Excel.FromCsv(manager.Memory);

            AssertRowsEqual(streamReader, memoryReader);
            AssertRowsEqual(streamReader, managerReader);
        }

        [Fact]
        public void XlsMemorySourceMatchesStreamSource()
        {
            byte[] bytes = XlsWorkbookBuilder.Build(sheets: [("S1", [["Name", 1, true]])]).ToArray();

            using XlsReader streamReader = Excel.FromXls(new MemoryStream(bytes, writable: false));
            using XlsReader memoryReader = Excel.FromXls(bytes.AsMemory());

            AssertRowsEqual(streamReader, memoryReader);
        }

        [Fact]
        public void XlsMemorySourceMatchesStreamSource_ForSlicedAndNonArrayBackedBuffers()
        {
            byte[] bytes = XlsWorkbookBuilder.Build(sheets: [("S1", [["Name", 1, true]])]).ToArray();
            byte[] prefixed = [0, 0, 0, .. bytes];
            ReadOnlyMemory<byte> sliced = prefixed.AsMemory(3, bytes.Length);
            var manager = new NonArrayMemoryManager(bytes);

            using XlsReader streamReader = Excel.FromXls(new MemoryStream(sliced.ToArray(), writable: false));
            using XlsReader memoryReader = Excel.FromXls(sliced);
            using XlsReader managerReader = Excel.FromXls(manager.Memory);

            AssertRowsEqual(streamReader, memoryReader);
            AssertRowsEqual(streamReader, managerReader);
        }

        [Fact]
        public void XlsMemorySourceThrowsForMalformedOleBuffers()
        {
            byte[] bytes = Encoding.UTF8.GetBytes("not an OLE document");

            Assert.Throws<InvalidDataException>(() => Excel.FromXls(bytes.AsMemory()));
        }

        private static void AssertRowsEqual(IExcelRowReader expected, IExcelRowReader actual)
        {
            string[] expectedValues = ReadRows(expected);
            string[] actualValues = ReadRows(actual);

            Assert.Equal(expectedValues, actualValues);
        }

        private static string[] ReadRows(IExcelRowReader reader)
        {
            List<string> values = [];
            foreach (Row row in reader)
            {
                StringBuilder sb = new();
                foreach (var cell in row.Cells)
                {
                    sb.Append(cell.Value.GetString());
                    sb.Append('|');
                }
                values.Add(sb.ToString());
            }
            return [.. values];
        }

        private sealed class NonArrayMemoryManager(byte[] data) : MemoryManager<byte>
        {
            private readonly byte[] _data = data;

            public override Memory<byte> Memory => _data;

            public override Span<byte> GetSpan()
            {
                throw new NotSupportedException();
            }

            public override MemoryHandle Pin(int elementIndex = 0)
            {
                throw new NotSupportedException();
            }

            public override void Unpin()
            {
                throw new NotSupportedException();
            }

#pragma warning disable IDISP010 // Call base.Dispose(disposing)
            protected override void Dispose(bool disposing)
#pragma warning restore IDISP010 // Call base.Dispose(disposing)
            {
            }
        }
    }
}
