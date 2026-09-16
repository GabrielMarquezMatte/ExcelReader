using BenchmarkDotNet.Attributes;
using ExcelReader.Native;

namespace ExcelReader.Benchmarks
{
    // NOTE: like NativeTypedParseBenchmark, MemoryDiagnoser sees MANAGED allocation only. The
    [MemoryDiagnoser]
    public class NativeRowReadBenchmark
    {
        private const int ExpectedRows = 65536;

        private byte[] _xlsb = [];
        private byte[] _buffer = new byte[64 * 1024];

        [GlobalSetup]
        public void Setup()
        {
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
                NativeApi.FreeRow(ref row);
                rows++;
            }
            return VerifyRows(rows);
        }

        [Benchmark]
        public int ReadAllBlob()
        {
            using NativeHandle handle = Open();
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
