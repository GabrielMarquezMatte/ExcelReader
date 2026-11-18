using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Writer;
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
        public void AReadingExportRejectingAnUnknownHandleDoesNotLeaveAnEarlierCallsError()
        {
            NativeApi.SetLastError("stale");
            byte* buffer = stackalloc byte[1];
            int written = Sentinel;

            int status = ((delegate* unmanaged<nint, byte*, int, int*, int>)&Exports.RowsNext)(0, buffer, 1, &written);

            Assert.Equal(NativeStatus.InvalidHandle, status);
            Assert.Empty(NativeApi.LastErrorText());
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

        [Fact]
        public void ParseTypedRejectingItsArgumentsStillZeroesTheTableOutParam()
        {
            NativeTable table;
            new Span<byte>(&table, sizeof(NativeTable)).Fill(0x5A);

            int status = ((delegate* unmanaged<nint, int, NativeColumnSpecRaw*, int, int, int, NativeTable*, int>)&Exports.ParseTyped)(
                0, 0, null, 0, 0, 0, &table);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, table.ColumnCount);
        }

        [Fact]
        public void ParseArrowRejectingItsArgumentsStillZeroesBothOutParams()
        {
            ArrowArray array;
            new Span<byte>(&array, sizeof(ArrowArray)).Fill(0x5A);

            int status = ((delegate* unmanaged<nint, int, NativeColumnSpecRaw*, int, int, int, ArrowArray*, ArrowSchema*, int>)&Exports.ParseArrow)(
                0, 0, null, 0, 0, 0, &array, null);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, array.Length);
        }

        [Fact]
        public void WriteTypedToMemoryRejectingItsArgumentsStillZeroesTheBufferOutParam()
        {
            NativeBuffer buffer = new() { Data = Sentinel, Length = Sentinel };

            int status = ((delegate* unmanaged<int, NativeColumnSpecRaw*, NativeTable*, NativeWriteOptionsRaw*, NativeBuffer*, int>)&Exports.WriteTypedToMemory)(
                0, null, null, null, &buffer);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(nint.Zero, buffer.Data);
            Assert.Equal(0, buffer.Length);
        }

        [Fact]
        public void EncryptPackageToMemoryRejectingItsArgumentsStillZeroesTheBufferOutParam()
        {
            NativeBuffer buffer = new() { Data = Sentinel, Length = Sentinel };

            int status = ((delegate* unmanaged<byte*, int, byte*, int, NativeBuffer*, int>)&Exports.EncryptPackageToMemory)(
                null, 0, null, 0, &buffer);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(nint.Zero, buffer.Data);
            Assert.Equal(0, buffer.Length);
        }

        [Fact]
        public void OpenWriteHandleRejectingItsArgumentsStillZeroesTheHandleOutParam()
        {
            nint outHandle = Sentinel;

            int status = ((delegate* unmanaged<byte*, int, int, NativeWriteOptionsRaw*, nint*, int>)&Exports.OpenWriteHandle)(
                null, 0, 0, null, &outHandle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, outHandle);
        }

        [Fact]
        public void AddStyleRejectingItsArgumentsStillZeroesTheStyleIdOutParam()
        {
            int outStyleId = Sentinel;

            int status = ((delegate* unmanaged<nint, NativeCellStyleRaw*, int*, int>)&Exports.AddStyle)(0, null, &outStyleId);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, outStyleId);
        }

        [Fact]
        public void RowsNextCellsRejectingItsArgumentsStillZeroesTheCountOutParam()
        {
            int outCount = Sentinel;

            int status = ((delegate* unmanaged<nint, NativeCellView**, int*, int>)&Exports.RowsNextCells)(0, null, &outCount);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Equal(0, outCount);
        }
    }
}
