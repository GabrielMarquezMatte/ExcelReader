using ExcelReader.Native;
using ExcelReader.Native.Csv;
using ExcelReader.Native.Typed;
using ExcelReader.Native.Writer;

namespace ExcelReader.Tests.Native
{
    public sealed unsafe class ExportsLastErrorTests
    {
        private static void AssertClearsStaleError(Func<int> call)
        {
            NativeApi.SetLastError("stale failure from an earlier call");
            int status = call();
            Assert.NotEqual(NativeStatus.Ok, status);
            Assert.DoesNotContain("stale failure", NativeApi.LastErrorText(), StringComparison.Ordinal);
        }

        [Fact]
        public void Every_Status_Export_Should_Clear_The_Last_Error_Of_An_Earlier_Call()
        {
            AssertClearsStaleError(() => ((delegate* unmanaged<byte*, int, int, NativeColumnSpecRaw*, NativeTable*, NativeWriteOptionsRaw*, int>)&Exports.WriteTyped)(null, 0, NativeFormat.Xlsx, null, null, null));
            AssertClearsStaleError(() => ((delegate* unmanaged<int, NativeColumnSpecRaw*, NativeTable*, NativeWriteOptionsRaw*, NativeBuffer*, int>)&Exports.WriteTypedToMemory)(NativeFormat.Xlsx, null, null, null, null));
            AssertClearsStaleError(() => ((delegate* unmanaged<byte*, int, int, NativeWriteOptionsRaw*, nint*, int>)&Exports.OpenWriteHandle)(null, 0, NativeFormat.Xlsx, null, null));
            AssertClearsStaleError(() => ((delegate* unmanaged<int, NativeWriteOptionsRaw*, nint*, int>)&Exports.OpenWriteHandleToMemory)(NativeFormat.Xlsx, null, null));
            AssertClearsStaleError(() => ((delegate* unmanaged<nint, NativeBuffer*, int>)&Exports.WriteHandleBytes)(0, null));
            AssertClearsStaleError(() => ((delegate* unmanaged<nint, int>)&Exports.StartRow)(0));
            AssertClearsStaleError(() => ((delegate* unmanaged<nint, long, int>)&Exports.WriteInt64)(0, 1));
            AssertClearsStaleError(() => ((delegate* unmanaged<nint, byte*, int, int>)&Exports.WriteString)(0, null, 0));
            AssertClearsStaleError(() => ((delegate* unmanaged<nint, int>)&Exports.CloseWriteHandle)(0));
            AssertClearsStaleError(() => ((delegate* unmanaged<byte*, int, byte*, int, byte*, int, int>)&Exports.EncryptPackage)(null, 0, null, 0, null, 0));
            AssertClearsStaleError(() => ((delegate* unmanaged<byte*, int, byte*, int, NativeBuffer*, int>)&Exports.EncryptPackageToMemory)(null, 0, null, 0, null));
            AssertClearsStaleError(() => ((delegate* unmanaged<byte*, int, NativeCsvAggregationRaw*, NativeCsvParallelOptionsRaw*, void**, int>)&Exports.CsvAggregateFile)(null, 0, null, null, null));
            AssertClearsStaleError(() => ((delegate* unmanaged<byte*, int, NativeCsvAggregationRaw*, NativeCsvParallelOptionsRaw*, void**, int>)&Exports.CsvAggregateMemory)(null, 0, null, null, null));
        }
    }
}
