namespace ExcelReader.CodSpeed
{
    // Loads the shared 65K-row dataset (linked from tests/ExcelReader.Benchmarks/Data) and builds the
    // synthetic record list used by the write scenarios.
    //
    // Every accessor is lazy on purpose: CodSpeed measures the whole process, so a scenario must only
    // pay for the fixture it actually reads. Loading happens once per process, before the measured
    // loop, so its cost is a constant offset instead of per-iteration work.
    internal static class Fixtures
    {
        private const string DataSetName = "65K_Records_Data";

        private static readonly string[] NamePool =
            ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];

        private static byte[]? _xlsx;
        private static byte[]? _xlsb;
        private static byte[]? _xls;
        private static byte[]? _csv;

        public static byte[] Xlsx => _xlsx ??= Load(".xlsx");

        public static byte[] Xlsb => _xlsb ??= Load(".xlsb");

        public static byte[] Xls => _xls ??= Load(".xls");

        public static byte[] Csv => _csv ??= Load(".csv");

        // `rows` records of [string, int, date, double] — the input every write scenario serializes.
        public static List<WriteRecord> Records(int rows)
        {
            List<WriteRecord> list = new(rows);
            for (int r = 1; r <= rows; r++)
            {
                list.Add(new WriteRecord
                {
                    Name = NamePool[r % NamePool.Length],
                    Id = r,
                    Date = DateTime.FromOADate(45292 + (r % 3650) + 0.25),
                    Value = r * 1.5,
                });
            }
            return list;
        }

        private static byte[] Load(string extension)
        {
            string path = Path.Combine(AppContext.BaseDirectory, "Data", DataSetName + extension);
            return File.ReadAllBytes(path);
        }
    }
}
