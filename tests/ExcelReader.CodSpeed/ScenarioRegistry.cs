namespace ExcelReader.CodSpeed
{
    // The scenario catalogue. Every entry here must have a matching benchmark in the repository-root
    // codspeed.yml; `--list` prints the names so the two can be diffed in CI.
    internal static class ScenarioRegistry
    {
        public static IReadOnlyList<Scenario> All { get; } =
        [
            // Reading the 65K-row, 14-column dataset cell by cell.
            Sync("read-xlsx", 5, Workloads.ReadXlsxStream),
            Sync("read-xlsx-memory", 5, Workloads.ReadXlsxMemory),
            Sync("read-xlsx-prefetch", 7, Workloads.ReadXlsxPrefetch),
            Sync("read-xlsx-materialized", 5, Workloads.ReadXlsxMaterialized),
            Awaited("read-xlsx-async", 5, Workloads.ReadXlsxAsync),
            Sync("read-xlsb", 7, Workloads.ReadXlsbStream),
            Sync("read-xls", 15, Workloads.ReadXlsStream),
            Sync("read-csv", 10, Workloads.ReadCsvStream),

            // Binding the same dataset to typed models.
            Sync("parse-xlsx-class", 5, Workloads.ParseXlsxClass),
            Sync("parse-xlsx-struct", 5, Workloads.ParseXlsxStruct),
            Sync("parse-csv-class", 7, Workloads.ParseCsvClass),

            // Serializing 50k records to each supported format.
            Write("write-xlsx", 10, records => Workloads.WriteXlsxAsync(records, useSharedStrings: false)),
            Write("write-xlsx-shared-strings", 12, records => Workloads.WriteXlsxAsync(records, useSharedStrings: true)),
            Write("write-xlsb", 25, Workloads.WriteXlsbAsync),
            Write("write-xls", 15, Workloads.WriteXlsAsync),
            Write("write-csv", 25, records => Task.FromResult(Workloads.WriteCsv(records))),
            Write("write-records-xlsx", 12, Workloads.WriteRecordsXlsxAsync),
        ];

        public static Scenario? Find(string name)
        {
            return All.FirstOrDefault(scenario => string.Equals(scenario.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        private static Scenario Sync(string name, int iterations, Func<long> body)
        {
            return new Scenario(name, iterations, count =>
            {
                long acc = 0;
                for (int i = 0; i < count; i++)
                {
                    acc += body();
                }
                return Task.FromResult(acc);
            });
        }

        private static Scenario Awaited(string name, int iterations, Func<Task<long>> body)
        {
            return new Scenario(name, iterations, async count =>
            {
                long acc = 0;
                for (int i = 0; i < count; i++)
                {
                    acc += await body();
                }
                return acc;
            });
        }

        // The record list is built once, outside the measured loop's per-iteration work, so the
        // scenario measures serialization rather than data generation.
        private static Scenario Write(string name, int iterations, Func<List<WriteRecord>, Task<long>> body)
        {
            return new Scenario(name, iterations, async count =>
            {
                List<WriteRecord> records = Fixtures.Records(Workloads.WriteRows);
                long acc = 0;
                for (int i = 0; i < count; i++)
                {
                    acc += await body(records);
                }
                return acc;
            });
        }
    }
}
