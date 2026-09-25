using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.NativePgoTrainer
{
    /// <summary>
    /// Runs the managed implementations of xl_parse_typed, xl_next_row and xl_next_row_view over the 65K
    /// benchmark fixtures, so a dotnet-trace of this process yields the static PGO profile NativeAOT compiles
    /// ExcelReader.Native with. See "Static PGO profile" in src/ExcelReader.Native/README.md.
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

        private static int Main()
        {
            string data = Path.Combine(AppContext.BaseDirectory, "data");
            (string File, int Format)[] fixtures =
            [
                ("65K_Records_Data.xlsx", NativeFormat.Xlsx),
                ("65K_Records_Data.xlsb", NativeFormat.Xlsb),
                ("65K_Records_Data.csv", NativeFormat.Csv),
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
                        status = TypedApi.ParseTyped(handle, Specs, headerRow: 1, out NativeTable table);
                        TypedApi.FreeTable(ref table);
                        ReadApi.Close(handle);
                    }
                    if (status == NativeStatus.Ok && (status = ReadApi.OpenMemory(bytes, format, out handle)) == NativeStatus.Ok)
                    {
                        while ((status = ReadApi.NextRow(handle, rowBuffer, out _)) == NativeStatus.Ok)
                        {
                        }
                        ReadApi.Close(handle);
                    }
                    if (status == NativeStatus.Eof && (status = ReadApi.OpenMemory(bytes, format, out handle)) == NativeStatus.Ok)
                    {
                        while ((status = ReadApi.NextRowView(handle, out _)) == NativeStatus.Ok)
                        {
                        }
                        ReadApi.Close(handle);
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

        private static NativeColumnSpec Column(string name, int type)
        {
            return new NativeColumnSpec { Names = [name], Type = type, Nullable = true };
        }
    }
}
