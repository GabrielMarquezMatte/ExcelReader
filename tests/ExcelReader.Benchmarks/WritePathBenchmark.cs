using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using Apache.Arrow;
using BenchmarkDotNet.Attributes;
using ExcelReader.Arrow;
using ExcelReader.Core.Writer;
using ExcelReader.Native;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    [SuppressMessage("Design", "CA1001", Justification = "BenchmarkDotNet owns the instance; GlobalCleanup releases everything.")]
    public unsafe class WritePathBenchmark
    {
        private const int Rows = 100_000;
        private const int StringColumns = 4;

        private readonly MemoryStream _output = new(32 * 1024 * 1024);
        private NativeColumn[] _columns = [];
        private readonly List<IntPtr> _native = [];
        private RecordBatch _batch = null!;

        [GlobalSetup]
        public void Setup()
        {
            _columns = new NativeColumn[StringColumns];
            IArrowArray[] arrays = new IArrowArray[StringColumns];
            Schema.Builder schema = new();
            for (int c = 0; c < StringColumns; c++)
            {
                int[] offsets = new int[Rows + 1];
                StringArray.Builder arrow = new();
                using MemoryStream data = new();
                for (int r = 0; r < Rows; r++)
                {
                    string value = $"customer-{c}-{r % 5000:D5}";
                    byte[] utf8 = Encoding.UTF8.GetBytes(value);
                    data.Write(utf8);
                    offsets[r + 1] = (int)data.Length;
                    arrow.Append(value);
                }
                _columns[c] = new NativeColumn
                {
                    Type = NativeColumnType.String,
                    Length = Rows,
                    Values = Pin(MemoryMarshal.AsBytes(offsets.AsSpan())),
                    Data = Pin(data.ToArray()),
                    DataLen = data.Length,
                };
                arrays[c] = arrow.Build();
                schema.Field(new Field($"c{c}", Apache.Arrow.Types.StringType.Default, nullable: false));
            }
            _batch = new RecordBatch(schema.Build(), arrays, Rows);
        }

        private IntPtr Pin(ReadOnlySpan<byte> bytes)
        {
            IntPtr p = Marshal.AllocHGlobal(bytes.Length);
            bytes.CopyTo(new Span<byte>((void*)p, bytes.Length));
            _native.Add(p);
            return p;
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            foreach (IntPtr p in _native)
            {
                Marshal.FreeHGlobal(p);
            }
            _native.Clear();
            _batch.Dispose();
            _output.Dispose();
        }

        [Benchmark]
        public long NativeStrings_Xlsx()
        {
            _output.SetLength(0);
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(_output, leaveOpen: true))
            {
                using XlsxSheetWriter sheet = workbook.AddSheet("S");
                sheet.Start();
                for (long r = 0; r < Rows; r++)
                {
                    using XlsxRowWriter row = sheet.StartRow();
                    foreach (NativeColumn column in _columns)
                    {
                        NativeApi.WriteCell(row, column, r);
                    }
                }
                sheet.End();
                workbook.End();
            }
            return _output.Length;
        }

        [Benchmark]
        public long NativeStrings_Csv()
        {
            _output.SetLength(0);
            using (CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(_output, leaveOpen: true))
            {
                using CsvSheetWriter sheet = workbook.AddSheet("S");
                sheet.Start();
                for (long r = 0; r < Rows; r++)
                {
                    using CsvRowWriter row = sheet.StartRow();
                    foreach (NativeColumn column in _columns)
                    {
                        NativeApi.WriteCell(row, column, r);
                    }
                }
                sheet.End();
                workbook.End();
            }
            return _output.Length;
        }

        [Benchmark]
        public long ArrowStrings_Xlsx()
        {
            _output.SetLength(0);
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(_output, leaveOpen: true))
            {
                workbook.WriteRecordBatch(_batch);
            }
            return _output.Length;
        }

        [Benchmark]
        public long StyledRows_Xlsx()
        {
            _output.SetLength(0);
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(_output, leaveOpen: true))
            {
                int style = workbook.AddStyle(new CellStyle { Bold = true });
                using XlsxSheetWriter sheet = workbook.AddSheet("S");
                sheet.Start();
                ISheetWriter<XlsxRowWriter> rows = sheet;
                for (int r = 0; r < Rows; r++)
                {
                    using XlsxRowWriter row = rows.StartRow(style);
                    row.Write(r);
                }
                sheet.End();
                workbook.End();
            }
            return _output.Length;
        }
    }
}
