using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    public sealed unsafe class CsvAggregateNativeTests
    {
        [Fact]
        public void Translate_Should_Return_Defaults_When_Options_Are_Absent()
        {
            int status = NativeCsvAggregateOptions.Translate(null, out CsvParallelOptions options);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.Equal(0, options.DegreeOfParallelism);
            Assert.Equal(0, options.HeaderRow);
            Assert.Equal((byte)',', options.Reader.Delimiter);
            Assert.Equal((byte)'"', options.Reader.Quote);
            Assert.True(options.Reader.DetectEncodingFromByteOrderMark);
            Assert.Equal(32 * 1024 * 1024, options.Reader.MaxCellBytes);
        }

        [Fact]
        public void Translate_Should_Apply_Dialect_When_Fields_Are_Set()
        {
            NativeCsvParallelOptionsRaw raw = new()
            {
                StructSize = sizeof(NativeCsvParallelOptionsRaw),
                DegreeOfParallelism = 4,
                HeaderRow = 1,
                Delimiter = ';',
                Quote = '\'',
                DetectBom = 1,
                MaxCellBytes = 4096,
            };

            int status = NativeCsvAggregateOptions.Translate(raw, out CsvParallelOptions options);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.Equal(4, options.DegreeOfParallelism);
            Assert.Equal(1, options.HeaderRow);
            Assert.Equal((byte)';', options.Reader.Delimiter);
            Assert.Equal((byte)'\'', options.Reader.Quote);
            Assert.False(options.Reader.DetectEncodingFromByteOrderMark);
            Assert.Equal(4096, options.Reader.MaxCellBytes);
        }

        [Theory]
        [InlineData(-1, 0, 0, 0, 0)]
        [InlineData(0, -1, 0, 0, 0)]
        [InlineData(0, 0, 256, 0, 0)]
        [InlineData(0, 0, 0, 256, 0)]
        [InlineData(0, 0, 0, 0, -1)]
        [InlineData(0, 0, 0, 0, 0, 3)]
        public void Translate_Should_Reject_When_A_Field_Is_Out_Of_Range(
            int dop, int headerRow, int delimiter, int quote, int maxCellBytes, int detectBom = 0)
        {
            NativeCsvParallelOptionsRaw raw = new()
            {
                StructSize = sizeof(NativeCsvParallelOptionsRaw),
                DegreeOfParallelism = dop,
                HeaderRow = headerRow,
                Delimiter = delimiter,
                Quote = quote,
                DetectBom = detectBom,
                MaxCellBytes = maxCellBytes,
            };

            int status = NativeCsvAggregateOptions.Translate(raw, out _);

            Assert.Equal(NativeStatus.InvalidArgument, status);
        }

        [Theory]
        [InlineData(4)]
        [InlineData(1024)]
        public void Translate_Should_Reject_When_StructSize_Is_Invalid(int structSize)
        {
            NativeCsvParallelOptionsRaw raw = new() { StructSize = structSize };

            int status = NativeCsvAggregateOptions.Translate(raw, out _);

            Assert.Equal(NativeStatus.InvalidArgument, status);
        }

        private static Row FirstRowOf(string csv)
        {
            CsvReader reader = Excel.FromCsv(Encoding.UTF8.GetBytes(csv), CsvReaderOptions.Default);
            CsvReader.Enumerator rows = reader.GetEnumerator();
            Assert.True(rows.MoveNext());
            return rows.Current;
        }

        [Fact]
        public void WriteRow_Should_Expose_Every_Field_As_A_NulTerminated_Cell()
        {
            using CsvAggregateState state = new();

            NativeRow row = state.WriteRow(FirstRowOf("alpha,beta,gamma\n"));

            Assert.Equal(3, row.CellCount);
            NativeRowCell* cells = (NativeRowCell*)row.Cells;
            Assert.Equal("alpha", ReadCell(cells[0]));
            Assert.Equal("beta", ReadCell(cells[1]));
            Assert.Equal("gamma", ReadCell(cells[2]));
            Assert.Equal(0, cells[0].Column);
            Assert.Equal(2, cells[2].Column);
            Assert.Equal(1, cells[0].Type);
            Assert.Equal(0, ((byte*)cells[0].Value)[cells[0].ValueLength]);
        }

        [Fact]
        public void WriteRow_Should_Report_Empty_Type_When_A_Field_Is_Blank()
        {
            using CsvAggregateState state = new();

            NativeRow row = state.WriteRow(FirstRowOf("alpha,,gamma\n"));

            NativeRowCell* cells = (NativeRowCell*)row.Cells;
            Assert.Equal(3, row.CellCount);
            Assert.Equal(0, cells[1].Type);
            Assert.Equal(0, cells[1].ValueLength);
        }

        [Fact]
        public void WriteRow_Should_Reuse_Its_Buffer_When_A_Later_Row_Is_Smaller()
        {
            using CsvAggregateState state = new();

            NativeRow wide = state.WriteRow(FirstRowOf("aaaaaaaaaa,bbbbbbbbbb,cccccccccc\n"));
            IntPtr firstBuffer = wide.Cells;
            NativeRow narrow = state.WriteRow(FirstRowOf("x\n"));

            Assert.Equal(firstBuffer, narrow.Cells);
            Assert.Equal(1, narrow.CellCount);
            Assert.Equal("x", ReadCell(*(NativeRowCell*)narrow.Cells));
        }

        private static string ReadCell(NativeRowCell cell)
        {
            return Encoding.UTF8.GetString((byte*)cell.Value, cell.ValueLength);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Tally
        {
            public int Seeds;
            public int Frees;
            public int Combines;
        }

        [UnmanagedCallersOnly]
        private static int SeedTally(void** outState, void* userData)
        {
            Tally* tally = (Tally*)userData;
            Interlocked.Increment(ref tally->Seeds);
            long* partial = (long*)NativeMemory.AllocZeroed((nuint)sizeof(long));
            *outState = partial;
            return 0;
        }

        [UnmanagedCallersOnly]
        private static int AccumulateFirstColumn(void* state, NativeRow* row, void* userData)
        {
            if (row->CellCount > 0)
            {
                NativeRowCell cell = *(NativeRowCell*)row->Cells;
                if (long.TryParse(
                        Encoding.UTF8.GetString((byte*)cell.Value, cell.ValueLength), CultureInfo.InvariantCulture, out long parsed))
                {
                    *(long*)state += parsed;
                }
            }
            return 0;
        }

        [UnmanagedCallersOnly]
        private static int CombineTally(void* accumulator, void* next, void* userData)
        {
            Tally* tally = (Tally*)userData;
            Interlocked.Increment(ref tally->Combines);
            *(long*)accumulator += *(long*)next;
            return 0;
        }

        [UnmanagedCallersOnly]
        private static void FreeTally(void* state, void* userData)
        {
            Tally* tally = (Tally*)userData;
            Interlocked.Increment(ref tally->Frees);
            NativeMemory.Free(state);
        }

        private static NativeCsvAggregationRaw TallyAggregation(Tally* tally)
        {
            return new NativeCsvAggregationRaw
            {
                StructSize = sizeof(NativeCsvAggregationRaw),
                Seed = (IntPtr)(delegate* unmanaged<void**, void*, int>)&SeedTally,
                Accumulate = (IntPtr)(delegate* unmanaged<void*, NativeRow*, void*, int>)&AccumulateFirstColumn,
                Combine = (IntPtr)(delegate* unmanaged<void*, void*, void*, int>)&CombineTally,
                FreeState = (IntPtr)(delegate* unmanaged<void*, void*, void>)&FreeTally,
                UserData = (IntPtr)tally,
            };
        }

        [UnmanagedCallersOnly]
        private static int SeedRecordThreadId(void** outState, void* userData)
        {
            *(int*)userData = Environment.CurrentManagedThreadId;
            long* partial = (long*)NativeMemory.AllocZeroed((nuint)sizeof(long));
            *outState = partial;
            return 0;
        }

        private static NativeCsvAggregationRaw ThreadRecordingAggregation(int* threadId)
        {
            return new NativeCsvAggregationRaw
            {
                StructSize = sizeof(NativeCsvAggregationRaw),
                Seed = (IntPtr)(delegate* unmanaged<void**, void*, int>)&SeedRecordThreadId,
                Accumulate = (IntPtr)(delegate* unmanaged<void*, NativeRow*, void*, int>)&AccumulateFirstColumn,
                Combine = (IntPtr)(delegate* unmanaged<void*, void*, void*, int>)&CombineTally,
                FreeState = (IntPtr)(delegate* unmanaged<void*, void*, void>)&FreeTally,
                UserData = (IntPtr)threadId,
            };
        }

        private static string WriteCsv(long rowCount)
        {
            string path = Path.Combine(Path.GetTempPath(), $"xlagg-{Guid.NewGuid():N}.csv");
            using StreamWriter writer = new(path);
            for (long index = 1; index <= rowCount; index++)
            {
                writer.WriteLine($"{index},filler-value-to-make-the-file-large-enough-to-partition");
            }
            return path;
        }

        [Fact]
        public void AggregateCsvFile_Should_Sum_Every_Row_And_Balance_Frees_When_Partitioned()
        {
            Tally tally = default;
            string path = WriteCsv(400_000);
            try
            {
                NativeCsvParallelOptionsRaw options = new()
                {
                    StructSize = sizeof(NativeCsvParallelOptionsRaw),
                    DegreeOfParallelism = 4,
                };

                int status = NativeApi.AggregateCsvFile(
                    Encoding.UTF8.GetBytes(path), TallyAggregation(&tally), options, out nint result);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(400_000L * 400_001L / 2, *(long*)result);
                Assert.True(tally.Seeds > 1);
                Assert.True(tally.Combines > 0);
                Assert.Equal(tally.Seeds - 1, tally.Frees);
                NativeMemory.Free((void*)result);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static string WriteQuotedFieldSpanningToEof()
        {
            string path = Path.Combine(Path.GetTempPath(), $"xlagg-{Guid.NewGuid():N}.csv");
            using StreamWriter writer = new(path);
            writer.Write("1,a\n2,\"");
            const string filler = "xxxxxxxxxx\n";
            const long target = 3 * 1024 * 1024;
            long written = 0;
            while (written < target)
            {
                writer.Write(filler);
                written += filler.Length;
            }
            writer.Write("\"\n");
            return path;
        }

        [Fact]
        public void AggregateCsvFile_Should_Free_Abandoned_Chunks_When_A_Quoted_Field_Spans_To_Eof()
        {
            Tally tally = default;
            string path = WriteQuotedFieldSpanningToEof();
            try
            {
                NativeCsvParallelOptionsRaw options = new()
                {
                    StructSize = sizeof(NativeCsvParallelOptionsRaw),
                    DegreeOfParallelism = 8,
                };

                int status = NativeApi.AggregateCsvFile(
                    Encoding.UTF8.GetBytes(path), TallyAggregation(&tally), options, out nint result);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(3L, *(long*)result);
                Assert.True(tally.Seeds > 1);
                Assert.Equal(0, tally.Combines);
                Assert.Equal(tally.Seeds - 1, tally.Frees);
                NativeMemory.Free((void*)result);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void AggregateCsvMemory_Should_Skip_The_Header_When_HeaderRow_Is_Set()
        {
            Tally tally = default;
            byte[] csv = Encoding.UTF8.GetBytes("amount,label\n10,a\n20,b\n30,c\n");
            fixed (byte* data = csv)
            {
                NativeCsvParallelOptionsRaw options = new()
                {
                    StructSize = sizeof(NativeCsvParallelOptionsRaw),
                    HeaderRow = 1,
                };

                int status = NativeApi.AggregateCsvMemory(
                    data, csv.Length, TallyAggregation(&tally), options, out nint result);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(60L, *(long*)result);
                Assert.Equal(tally.Seeds - 1, tally.Frees);
                NativeMemory.Free((void*)result);
            }
        }

        [Fact]
        public void AggregateCsvMemory_Should_Return_A_Seeded_State_When_The_Source_Is_Empty()
        {
            Tally tally = default;
            byte[] csv = [];
            fixed (byte* data = csv)
            {
                int status = NativeApi.AggregateCsvMemory(
                    data, 0, TallyAggregation(&tally), null, out nint result);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(0L, *(long*)result);
                Assert.Equal(1, tally.Seeds);
                Assert.Equal(0, tally.Frees);
                Assert.Equal(0, tally.Combines);
                NativeMemory.Free((void*)result);
            }
        }

        [Fact]
        public void AggregateCsvMemory_Should_Not_Combine_When_DegreeOfParallelism_Is_One()
        {
            Tally tally = default;
            byte[] csv = Encoding.UTF8.GetBytes("1,a\n2,b\n3,c\n");
            fixed (byte* data = csv)
            {
                NativeCsvParallelOptionsRaw options = new()
                {
                    StructSize = sizeof(NativeCsvParallelOptionsRaw),
                    DegreeOfParallelism = 1,
                };

                int status = NativeApi.AggregateCsvMemory(
                    data, csv.Length, TallyAggregation(&tally), options, out nint result);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(6L, *(long*)result);
                Assert.Equal(1, tally.Seeds);
                Assert.Equal(0, tally.Combines);
                Assert.Equal(0, tally.Frees);
                NativeMemory.Free((void*)result);
            }
        }

        [Fact]
        public void AggregateCsvMemory_Should_Not_Run_Seed_On_The_Calling_Thread_When_DegreeOfParallelism_Is_One()
        {
            int seedThreadId = -1;
            byte[] csv = Encoding.UTF8.GetBytes("1,a\n2,b\n3,c\n");
            fixed (byte* data = csv)
            {
                NativeCsvParallelOptionsRaw options = new()
                {
                    StructSize = sizeof(NativeCsvParallelOptionsRaw),
                    DegreeOfParallelism = 1,
                };

                int status = NativeApi.AggregateCsvMemory(
                    data, csv.Length, ThreadRecordingAggregation(&seedThreadId), options, out nint result);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.NotEqual(Environment.CurrentManagedThreadId, seedThreadId);
                NativeMemory.Free((void*)result);
            }
        }
    }
}
