using ExcelReader.Native;
using ExcelReader.Native.Csv;

namespace ExcelReader.Tests.Native
{
    public sealed unsafe class ExportsOutParamTests
    {
        private const int Sentinel = 0x5A5A5A5A;

        public static TheoryData<string> BufferExports => ["xl_rows_next", "xl_rows_read_all_blob", "xl_last_error"];

        private static int CallBufferExport(string export, int capacity, int* outValue)
        {
            return export switch
            {
                "xl_rows_next" => ((delegate* unmanaged<nint, byte*, int, int*, int>)&Exports.RowsNext)(0, null, capacity, outValue),
                "xl_rows_read_all_blob" => ((delegate* unmanaged<nint, byte*, int, int*, int>)&Exports.RowsReadAllBlob)(0, null, capacity, outValue),
                "xl_last_error" => ((delegate* unmanaged<byte*, int, int*, int>)&Exports.LastError)(null, capacity, outValue),
                _ => throw new ArgumentOutOfRangeException(nameof(export)),
            };
        }

        [Theory]
        [MemberData(nameof(BufferExports))]
        public void ABufferExportRejectingItsArgumentsStillZeroesTheLengthOutParam(string export)
        {
            int outValue = Sentinel;

            int status = CallBufferExport(export, capacity: -1, &outValue);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, outValue);
        }

        [Fact]
        public void SheetNameAtRejectingItsArgumentsStillZeroesTheLengthOutParam()
        {
            int outValue = Sentinel;

            int status = ((delegate* unmanaged<nint, int, byte*, int, int*, int>)&Exports.SheetNameAt)(0, 0, null, -1, &outValue);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, outValue);
        }

        [Fact]
        public void CsvAggregateMemoryRejectingItsArgumentsStillNullsTheStateOutParam()
        {
            void* outState = (void*)Sentinel;

            int status = ((delegate* unmanaged<byte*, int, NativeCsvAggregationRaw*, NativeCsvParallelOptionsRaw*, void**, int>)&Exports.CsvAggregateMemory)(
                null, 0, null, null, &outState);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(nint.Zero, (nint)outState);
        }

        [Fact]
        public void CsvAggregateFileRejectingItsArgumentsStillNullsTheStateOutParam()
        {
            void* outState = (void*)Sentinel;

            int status = ((delegate* unmanaged<byte*, int, NativeCsvAggregationRaw*, NativeCsvParallelOptionsRaw*, void**, int>)&Exports.CsvAggregateFile)(
                null, 0, null, null, &outState);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(nint.Zero, (nint)outState);
        }

        [Fact]
        public void SheetIndexRejectingItsArgumentsStillWritesMinusOne()
        {
            int outValue = Sentinel;

            int status = ((delegate* unmanaged<nint, byte*, int, int*, int>)&Exports.SheetIndex)(0, null, 5, &outValue);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(-1, outValue);
        }

        [Fact]
        public void RowsOpenOnAnUnknownWorkbookStillZeroesTheCursorOutParam()
        {
            nint outRows = Sentinel;

            int status = ((delegate* unmanaged<nint, int, nint*, int>)&Exports.RowsOpen)(0, 0, &outRows);

            Assert.Equal(NativeStatus.InvalidHandle, status);
            Assert.Equal(0, outRows);
        }
    }
}
