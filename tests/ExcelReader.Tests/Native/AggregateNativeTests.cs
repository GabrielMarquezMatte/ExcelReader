using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;
using ExcelReader.Native;
using ExcelReader.Native.Csv;
using ExcelReader.Native.Reading;

namespace ExcelReader.Tests.Native
{
    public sealed unsafe class AggregateNativeTests
    {
        private const int Rows = 60_000;

        [StructLayout(LayoutKind.Sequential)]
        private struct Partial
        {
            public long Sum;
            public long Rows;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Tally
        {
            public int Seeds;
            public int Frees;
            public int Combines;
            public int FailAt;
            public int Seen;
        }

        private static void WriteRows<TSheet, TRow>(IWorkbookWriter<TSheet> workbook)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using TSheet sheet = workbook.AddSheet("S1");
            using (TRow header = sheet.StartRow())
            {
                header.Write("id");
                header.Write("label");
            }
            for (long index = 1; index <= Rows; index++)
            {
                using TRow row = sheet.StartRow();
                row.Write(index);
                row.Write("filler-value-" + index.ToString(CultureInfo.InvariantCulture));
            }
            sheet.End();
            workbook.End();
        }

        private static byte[] Build(int format)
        {
            using MemoryStream stream = new();
            switch (format)
            {
                case NativeFormat.Xlsx:
                    using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true))
                    {
                        WriteRows<XlsxSheetWriter, XlsxRowWriter>(workbook);
                    }
                    break;
                case NativeFormat.Xlsb:
                    using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(stream, leaveOpen: true))
                    {
                        WriteRows<XlsbSheetWriter, XlsbRowWriter>(workbook);
                    }
                    break;
                case NativeFormat.Xls:
                    using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(stream, leaveOpen: true))
                    {
                        WriteRows<XlsSheetWriter, XlsRowWriter>(workbook);
                    }
                    break;
                default:
                    using (CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(stream, leaveOpen: true))
                    {
                        WriteRows<CsvSheetWriter, CsvRowWriter>(workbook);
                    }
                    break;
            }
            return stream.ToArray();
        }

        [UnmanagedCallersOnly]
        private static int Seed(void** outState, void* userData)
        {
            Interlocked.Increment(ref ((Tally*)userData)->Seeds);
            *outState = NativeMemory.AllocZeroed((nuint)sizeof(Partial));
            return 0;
        }

        [UnmanagedCallersOnly]
        private static int SumFirstColumn(void* state, NativeCellView* cells, int count, void* userData)
        {
            Tally* tally = (Tally*)userData;
            if (tally->FailAt != 0 && Interlocked.Increment(ref tally->Seen) >= tally->FailAt)
            {
                return 7;
            }
            Partial* partial = (Partial*)state;
            partial->Rows++;
            if (count > 0 && cells[0].Column == 0)
            {
                NativeCellView cell = cells[0];
                partial->Sum += cell.HasNumber != 0 && cell.ValueLength == 0
                    ? (long)cell.Number
                    : long.Parse(Encoding.UTF8.GetString(cell.Value, cell.ValueLength), CultureInfo.InvariantCulture);
            }
            return 0;
        }

        [UnmanagedCallersOnly]
        private static int Combine(void* accumulator, void* next, void* userData)
        {
            Interlocked.Increment(ref ((Tally*)userData)->Combines);
            ((Partial*)accumulator)->Sum += ((Partial*)next)->Sum;
            ((Partial*)accumulator)->Rows += ((Partial*)next)->Rows;
            return 0;
        }

        [UnmanagedCallersOnly]
        private static void Free(void* state, void* userData)
        {
            Interlocked.Increment(ref ((Tally*)userData)->Frees);
            NativeMemory.Free(state);
        }

        private static NativeCsvAggregationRaw Aggregation(Tally* tally)
        {
            return new NativeCsvAggregationRaw
            {
                StructSize = sizeof(NativeCsvAggregationRaw),
                Seed = (IntPtr)(delegate* unmanaged<void**, void*, int>)&Seed,
                Accumulate = (IntPtr)(delegate* unmanaged<void*, NativeCellView*, int, void*, int>)&SumFirstColumn,
                Combine = (IntPtr)(delegate* unmanaged<void*, void*, void*, int>)&Combine,
                FreeState = (IntPtr)(delegate* unmanaged<void*, void*, void>)&Free,
                UserData = (IntPtr)tally,
            };
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, true)]
        [InlineData(NativeFormat.Xlsb, true)]
        [InlineData(NativeFormat.Xls, false)]
        [InlineData(NativeFormat.Csv, false)]
        public void AggregateWorkbook_Should_Sum_Every_Data_Row_Once_And_Balance_Frees(int format, bool partitions)
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(Build(format), format, out NativeHandle? handle));
            try
            {
                Tally tally = default;
                int status = CsvAggregateApi.AggregateWorkbook(handle, 0, Aggregation(&tally), 1, 4, out nint result);

                Assert.Equal(NativeStatus.Ok, status);
                Partial* total = (Partial*)result;
                Assert.Equal((long)Rows, total->Rows);
                Assert.Equal((long)Rows * (Rows + 1) / 2, total->Sum);
                Assert.Equal(tally.Seeds - 1, tally.Frees);
                if (partitions)
                {
                    Assert.True(tally.Combines > 0, "the fixture must partition for this format");
                }
                NativeMemory.Free(total);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void AggregateWorkbook_Should_Return_The_Callers_Code_And_Free_Every_State()
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(Build(NativeFormat.Xlsx), NativeFormat.Xlsx, out NativeHandle? handle));
            try
            {
                Tally tally = new() { FailAt = 20_000 };
                int status = CsvAggregateApi.AggregateWorkbook(handle, 0, Aggregation(&tally), 1, 4, out nint result);

                Assert.Equal(7, status);
                Assert.Equal(0, result);
                Assert.Equal(tally.Seeds, tally.Frees);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void AggregateWorkbook_Should_Reject_A_Negative_Option()
        {
            Assert.Equal(NativeStatus.Ok, ReadApi.OpenMemory(Build(NativeFormat.Csv), NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                Tally tally = default;
                Assert.Equal(NativeStatus.InvalidArgument, CsvAggregateApi.AggregateWorkbook(handle, 0, Aggregation(&tally), -1, 0, out _));
                Assert.Equal(NativeStatus.InvalidArgument, CsvAggregateApi.AggregateWorkbook(handle, 0, Aggregation(&tally), 0, -1, out _));
                Assert.Equal(0, tally.Seeds);
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [Fact]
        public void AggregateWorkbook_Should_Reject_A_Null_Handle()
        {
            Tally tally = default;
            Assert.Equal(NativeStatus.InvalidHandle, CsvAggregateApi.AggregateWorkbook(null, 0, Aggregation(&tally), 0, 0, out _));
        }

        [Fact]
        public void Aggregate_Export_Should_Reject_A_Null_Aggregation()
        {
            void* state;
            int status = ((delegate* unmanaged<nint, int, NativeCsvAggregationRaw*, int, int, void**, int>)&Exports.Aggregate)(0, 0, null, 0, 0, &state);
            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Contains("aggregation must not be NULL", NativeApi.LastErrorText(), StringComparison.Ordinal);
        }
    }
}
