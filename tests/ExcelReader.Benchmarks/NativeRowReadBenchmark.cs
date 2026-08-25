using BenchmarkDotNet.Attributes;
using ExcelReader.Native;

namespace ExcelReader.Benchmarks
{
    // The row-oriented half of the C ABI, which NativeTypedParseBenchmark does not touch: the four
    // ways a caller can pull raw rows out of a sheet, all over the same fixture so their costs can
    // be read against each other and against ParseTyped's columnar number.
    //
    //   NextRowBlob      xl_next_row          — one crossing per row, caller-owned buffer
    //   NextRowDecoded   (internal only)      — one crossing per row, one native block per row
    //                                            ReadAllDecoded's per-row primitive, not itself an
    //                                            exported ABI function
    //   ReadAllBlob      xl_read_all_blob     — one crossing for the sheet, caller-owned buffer
    //   ReadAllDecoded   xl_read_all_decoded  — one crossing for the sheet, one native block per row
    //
    // Drives NativeApi rather than the C ABI for the reason ExcelReader.Native's csproj states:
    // [UnmanagedCallersOnly] exports cannot be invoked from managed code. That means the per-row
    // ones do NOT pay a real native-to-managed transition here, so the gap between the per-row and
    // whole-sheet pairs is a floor, not the figure a C caller sees.
    //
    // NOTE: like NativeTypedParseBenchmark, MemoryDiagnoser sees MANAGED allocation only. The
    // decoded paths' Marshal.AllocHGlobal blocks are the output, not the overhead, and are
    // invisible here — wall clock is what moves when those paths change.
    [MemoryDiagnoser]
    public class NativeRowReadBenchmark
    {
        // Every row of the fixture, header included: none of these paths knows about a header row.
        private const int ExpectedRows = 65536;

        private byte[] _xlsb = [];
        private byte[] _buffer = new byte[64 * 1024];

        [GlobalSetup]
        public void Setup()
        {
            // Loud failure if the fixture was never copied to the output directory — a benchmark
            // that silently ran on nothing would publish a number that reads like a large win.
            _xlsb = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Data", "65K_Records_Data.xlsb"));
        }

        [Benchmark(Baseline = true)]
        public int NextRowBlob()
        {
            using NativeHandle handle = Open();
            int rows = 0;
            while (true)
            {
                int status = NativeApi.NextRow(handle, _buffer, out int written);
                if (status == NativeStatus.BufferTooSmall)
                {
                    // The row stays pending on the handle, so the retry costs a copy, not a re-read.
                    _buffer = new byte[written];
                    status = NativeApi.NextRow(handle, _buffer, out written);
                }
                if (status == NativeStatus.Eof)
                {
                    break;
                }
                Verify(status);
                rows++;
            }
            return VerifyRows(rows);
        }

        [Benchmark]
        public int NextRowDecoded()
        {
            using NativeHandle handle = Open();
            int rows = 0;
            while (true)
            {
                int status = NativeApi.NextRowDecoded(handle, out NativeRow row);
                if (status == NativeStatus.Eof)
                {
                    break;
                }
                Verify(status);
                // Freed per row, as a real caller must: holding all 65K would measure a workload
                // nobody runs, and the free is part of this path's cost.
                NativeApi.FreeRow(ref row);
                rows++;
            }
            return VerifyRows(rows);
        }

        [Benchmark]
        public int ReadAllBlob()
        {
            using NativeHandle handle = Open();
            // The ask-the-size call is where the whole sheet is actually read and accumulated; the
            // second call is the copy out. Both are part of what a caller pays.
            int status = NativeApi.ReadAllBlob(handle, Span<byte>.Empty, out int written);
            if (status != NativeStatus.BufferTooSmall)
            {
                throw new InvalidOperationException($"expected XL_BUFFER_TOO_SMALL sizing probe, got status {status}.");
            }
            if (_buffer.Length < written)
            {
                _buffer = new byte[written];
            }
            Verify(NativeApi.ReadAllBlob(handle, _buffer.AsSpan(0, written), out _));
            return VerifyRows(RowCountOf(_buffer));
        }

        [Benchmark]
        public int ReadAllDecoded()
        {
            using NativeHandle handle = Open();
            int status = NativeApi.ReadAllDecoded(handle, out NativeRows rows);
            try
            {
                Verify(status);
                return VerifyRows(rows.RowCount);
            }
            finally
            {
                NativeApi.FreeRows(ref rows);
            }
        }

        // The blob opens with an int32 row count (see excelreader.h on xl_read_all_blob).
        private static int RowCountOf(byte[] blob)
        {
            return System.Buffers.Binary.BinaryPrimitives.ReadInt32LittleEndian(blob);
        }

        private NativeHandle Open()
        {
            int status = NativeApi.OpenMemory(_xlsb, NativeFormat.Xlsb, out NativeHandle? handle);
            if (status != NativeStatus.Ok || handle is null)
            {
                throw new InvalidOperationException($"opening the xlsb fixture failed with status {status}.");
            }
            return handle;
        }

        private static void Verify(int status)
        {
            if (status != NativeStatus.Ok)
            {
                throw new InvalidOperationException($"native row read failed with status {status}.");
            }
        }

        // Checked every iteration, not once in setup: a read that quietly yields nothing still
        // publishes a timing, and a timing measured over no work reads exactly like a win.
        private static int VerifyRows(int rows)
        {
            if (rows != ExpectedRows)
            {
                throw new InvalidOperationException($"expected {ExpectedRows} rows, got {rows}.");
            }
            return rows;
        }
    }
}
