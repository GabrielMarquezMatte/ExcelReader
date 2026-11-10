using System.Runtime.InteropServices;
using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;
using ExcelReader.Native.Writer;

namespace ExcelReader.NativePgoTrainer
{
    /// <summary>
    /// Runs the managed implementations of xl_parse_typed, xl_parse_arrow (dictionary strings), xl_rows_next, xl_rows_next_view, xl_rows_next_cells and
    /// xl_rows_foreach over the 65K benchmark fixtures, so a dotnet-trace of this process yields the static PGO
    /// profile NativeAOT compiles ExcelReader.Native with. See "Static PGO profile" in src/ExcelReader.Native/README.md.
    /// </summary>
    internal static class Program
    {
        private const int Iterations = 30;

        private static readonly NativeColumnSpec[] Specs =
        [
            Column("Region", NativeColumnType.String),
            Column("Country", NativeColumnType.String),
            Column("Item Type", NativeColumnType.String),
            Column("Sales Channel", NativeColumnType.String),
            Column("Order Priority", NativeColumnType.String),
            Column("Order Date", NativeColumnType.Date),
            Column("Order ID", NativeColumnType.Int64),
            Column("Ship Date", NativeColumnType.Date),
            Column("Units Sold", NativeColumnType.Int64),
            Column("Unit Price", NativeColumnType.Float64),
            Column("Unit Cost", NativeColumnType.Float64),
            Column("Total Revenue", NativeColumnType.Float64),
            Column("Total Cost", NativeColumnType.Float64),
            Column("Total Profit", NativeColumnType.Float64),
        ];

        private static readonly NativeColumnSpec[] DictionarySpecs = Array.ConvertAll(Specs,
            static spec => spec.Type == NativeColumnType.String ? spec with { Type = NativeColumnType.StringDictionary } : spec);

        private static int Main()
        {
            string data = Path.Combine(AppContext.BaseDirectory, "data");
            (string File, int Format)[] fixtures =
            [
                ("65K_Records_Data.xlsx", NativeFormat.Xlsx),
                ("65K_Records_Data.xlsb", NativeFormat.Xlsb),
                ("65K_Records_Data.csv", NativeFormat.Csv),
                ("65K_Records_Data.xls", NativeFormat.Xls),
            ];
            byte[] rowBuffer = new byte[1 << 16];
            foreach ((string file, int format) in fixtures)
            {
                byte[] bytes = File.ReadAllBytes(Path.Combine(data, file));
                for (int i = 0; i < Iterations; i++)
                {
                    int status = ReadApi.OpenMemory(bytes, format, out NativeHandle? handle);
                    if (status == NativeStatus.Ok)
                    {
                        status = TypedApi.ParseTyped(handle, 0, Specs, headerRow: 1, out NativeTable table);
                        if (status == NativeStatus.Ok)
                        {
                            status = WriteAll(table);
                        }
                        TypedApi.FreeTable(ref table);
                        if (status == NativeStatus.Ok)
                        {
                            status = ParseDictionaryArrow(handle);
                        }
                        if (status == NativeStatus.Ok && format != NativeFormat.Csv)
                        {
                            status = TypedApi.ParseTypedTable(handle, 0, Specs, headerRow: 1, degreeOfParallelism: 0, out table);
                            TypedApi.FreeTable(ref table);
                        }
                        ReadApi.Close(handle);
                    }
                    if (status == NativeStatus.Ok && (status = ReadApi.OpenMemory(bytes, format, out handle)) == NativeStatus.Ok)
                    {
                        ReadApi.OpenRows(handle, 0, out NativeRowCursor? rows);
                        using NativeRowCursor cursor = rows!;
                        while ((status = ReadApi.NextRow(cursor, rowBuffer, out _)) == NativeStatus.Ok)
                        {
                        }
                        ReadApi.Close(handle);
                    }
                    if (status == NativeStatus.Eof && (status = ReadApi.OpenMemory(bytes, format, out handle)) == NativeStatus.Ok)
                    {
                        ReadApi.OpenRows(handle, 0, out NativeRowCursor? viewRows);
                        using NativeRowCursor viewCursor = viewRows!;
                        while ((status = ReadApi.NextRowView(viewCursor, out _)) == NativeStatus.Ok)
                        {
                        }
                        ReadApi.Close(handle);
                    }
                    if (status == NativeStatus.Eof)
                    {
                        status = ReadCellsAndForEach(bytes, format);
                    }
                    if (status != NativeStatus.Eof)
                    {
                        Console.Error.WriteLine($"{file}: failed with status {status}.");
                        return 1;
                    }
                }
            }
            return 0;
        }

        private static unsafe int ParseDictionaryArrow(NativeHandle? handle)
        {
            int status = ArrowApi.ParseArrow(handle, 0, DictionarySpecs, headerRow: 1, out ArrowArray array, out ArrowSchema schema);
            ArrowApi.ReleaseArrowArray((IntPtr)(&array));
            ArrowApi.ReleaseArrowSchema((IntPtr)(&schema));
            return status;
        }

        private static unsafe int ReadCellsAndForEach(byte[] bytes, int format)
        {
            int status = ReadApi.OpenMemory(bytes, format, out NativeHandle? handle);
            if (status != NativeStatus.Ok)
            {
                return status;
            }
            try
            {
                ReadApi.OpenRows(handle, 0, out NativeRowCursor? rows);
                using (NativeRowCursor cursor = rows!)
                {
                    while ((status = ReadApi.NextRowCells(cursor, out _, out _)) == NativeStatus.Ok)
                    {
                    }
                }
                if (status != NativeStatus.Eof)
                {
                    return status;
                }
                status = ReadApi.ForEachRow(handle, 0, &SkipRow, null);
                return status == NativeStatus.Ok ? NativeStatus.Eof : status;
            }
            finally
            {
                ReadApi.Close(handle);
            }
        }

        [UnmanagedCallersOnly]
        private static unsafe int SkipRow(NativeCellView* cells, int count, void* userData)
        {
            return NativeStatus.Ok;
        }

        // The parsed table written back out in every format xl_write_typed takes, serially and in parallel.
        private static int WriteAll(NativeTable table)
        {
            foreach (int format in (int[])[NativeFormat.Xlsx, NativeFormat.Xlsb, NativeFormat.Csv])
            {
                foreach (int degree in (int[])[1, 0])
                {
                    int status = WriteApi.WriteTypedToMemory(format, Specs, table, new NativeWriteOptions { DegreeOfParallelism = degree }, out _);
                    if (status != NativeStatus.Ok)
                    {
                        return status;
                    }
                }
            }
            return NativeStatus.Ok;
        }

        private static NativeColumnSpec Column(string name, int type)
        {
            return new NativeColumnSpec { Names = [name], Type = type, Nullable = true };
        }
    }
}
