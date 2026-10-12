using System.Runtime.InteropServices;
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelReader.Native;
using ExcelReader.Native.Reading;

namespace ExcelReader.Benchmarks
{
    // NOTE: the xl_source/xl_stream callbacks are [UnmanagedCallersOnly] methods called back into this
    // process, so each one pays a managed->unmanaged->managed transition a native caller would not.
    [MemoryDiagnoser]
    public unsafe class NativeOpenBenchmark
    {
        private sealed class StreamState(byte[] bytes)
        {
            internal readonly byte[] Bytes = bytes;
            internal int Position;
        }

        private byte[] _data = [];
        private byte[] _utf8Path = [];
        private int _format;
        private int _expectedRows;

        [Params("csv", "xlsx", "xlsb", "xls")]
        public string Format { get; set; } = "";

        [Params(false, true)]
        public bool ReadRows { get; set; }

        [GlobalSetup]
        public void Setup()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "Data", $"65K_Records_Data.{Format}");
            _data = File.ReadAllBytes(path);
            _utf8Path = Encoding.UTF8.GetBytes(path);
            _format = Format switch
            {
                "csv" => NativeFormat.Csv,
                "xlsx" => NativeFormat.Xlsx,
                "xlsb" => NativeFormat.Xlsb,
                _ => NativeFormat.Xls,
            };
            Verify(ReadApi.OpenMemory(_data, _format, out NativeHandle? handle), handle);
            using (handle)
            {
                _expectedRows = CountRows(handle!);
            }
        }

        [Benchmark(Baseline = true)]
        public int OpenMemory()
        {
            return Consume(ReadApi.OpenMemory(_data, _format, out NativeHandle? handle), handle);
        }

        [Benchmark]
        public int OpenFile()
        {
            return Consume(ReadApi.OpenFile(_utf8Path, _format, out NativeHandle? handle), handle);
        }

        [Benchmark]
        public int OpenSource()
        {
            NativeSourceRaw raw = new()
            {
                StructSize = sizeof(NativeSourceRaw),
                UserData = (void*)GCHandle.ToIntPtr(GCHandle.Alloc(_data)),
                Length = _data.Length,
                ReadAt = &ReadAt,
                Release = &Release,
            };
            if (!CallbackByteSource.TryCreate(&raw, hasOutHandle: true, out CallbackByteSource? source, out string? error))
            {
                throw new InvalidOperationException(error);
            }
            return Consume(ReadApi.OpenSource(source, _format, null, out NativeHandle? handle), handle);
        }

        [Benchmark]
        public int OpenStream()
        {
            NativeStreamRaw raw = new()
            {
                StructSize = sizeof(NativeStreamRaw),
                UserData = (void*)GCHandle.ToIntPtr(GCHandle.Alloc(new StreamState(_data))),
                Read = &Read,
                Release = &Release,
            };
            if (!CallbackReadStream.TryCreate(&raw, hasOutHandle: true, out CallbackReadStream? stream, out string? error))
            {
                throw new InvalidOperationException(error);
            }
            return Consume(ReadApi.OpenStream(stream, _format, null, out NativeHandle? handle), handle);
        }

        [UnmanagedCallersOnly]
        private static long ReadAt(void* userData, long offset, byte* buffer, long length)
        {
            byte[] bytes = (byte[])GCHandle.FromIntPtr((nint)userData).Target!;
            int count = (int)Math.Min(length, bytes.Length - offset);
            bytes.AsSpan((int)offset, count).CopyTo(new Span<byte>(buffer, count));
            return count;
        }

        [UnmanagedCallersOnly]
        private static long Read(void* userData, byte* buffer, long length)
        {
            StreamState state = (StreamState)GCHandle.FromIntPtr((nint)userData).Target!;
            int count = (int)Math.Min(length, state.Bytes.Length - state.Position);
            state.Bytes.AsSpan(state.Position, count).CopyTo(new Span<byte>(buffer, count));
            state.Position += count;
            return count;
        }

        [UnmanagedCallersOnly]
        private static void Release(void* userData)
        {
            GCHandle.FromIntPtr((nint)userData).Free();
        }

        private int Consume(int status, NativeHandle? handle)
        {
            Verify(status, handle);
            using (handle)
            {
                if (!ReadRows)
                {
                    return 0;
                }
                int rows = CountRows(handle!);
                if (rows != _expectedRows)
                {
                    throw new InvalidOperationException($"expected {_expectedRows} rows, got {rows}.");
                }
                return rows;
            }
        }

        private static int CountRows(NativeHandle handle)
        {
            ReadApi.OpenRows(handle, 0, out NativeRowCursor? opened);
            using NativeRowCursor cursor = opened!;
            int rows = 0;
            int status;
            while ((status = ReadApi.NextRowView(cursor, out _)) == NativeStatus.Ok)
            {
                rows++;
            }
            if (status != NativeStatus.Eof)
            {
                throw new InvalidOperationException($"native row read failed with status {status}: {NativeApi.LastErrorText()}");
            }
            return rows;
        }

        private static void Verify(int status, NativeHandle? handle)
        {
            if (status != NativeStatus.Ok || handle is null)
            {
                throw new InvalidOperationException($"open failed with status {status}: {NativeApi.LastErrorText()}");
            }
        }
    }
}
