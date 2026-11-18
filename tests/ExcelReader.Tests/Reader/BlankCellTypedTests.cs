using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Columnar;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Tests.Reader
{
    public sealed class BlankCellTypedTests
    {
        private static readonly DateTime Sample = new(2024, 3, 5, 12, 0, 0);

        private sealed class BlankRow
        {
            public int? Qty { get; set; }
            public double? Rate { get; set; }
            public DateTime? When { get; set; }
            public DateOnly? Day { get; set; }
            public decimal? Amount { get; set; }
            public bool? Flag { get; set; }
        }

        private readonly struct AcceptsEmpty : IUtf8SpanParsable<AcceptsEmpty>
        {
            public static AcceptsEmpty Parse(ReadOnlySpan<byte> utf8Text, IFormatProvider? provider)
            {
                return default;
            }

            public static bool TryParse(ReadOnlySpan<byte> utf8Text, IFormatProvider? provider, out AcceptsEmpty result)
            {
                result = default;
                return true;
            }
        }

        public static TheoryData<ExcelColumnType, bool, bool, int> Cases()
        {
            TheoryData<ExcelColumnType, bool, bool, int> data = [];
            foreach (ExcelColumnType type in Enum.GetValues<ExcelColumnType>())
            {
                foreach (bool nullable in new[] { false, true })
                {
                    foreach (bool csv in new[] { false, true })
                    {
                        data.Add(type, nullable, csv, 1);
                    }
                }
            }
            data.Add(ExcelColumnType.Int64Column, true, true, 4);
            data.Add(ExcelColumnType.TimestampColumn, true, false, 4);
            return data;
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void Blank_And_Missing_Cells_Append_Per_Column_Nullability(ExcelColumnType type, bool nullable, bool csv, int dop)
        {
            using IExcelWorkbook workbook = csv ? OpenCsv(type) : OpenXlsx(type);
            TypedColumnSpec[] specs =
            [
                new() { Names = ["k"], Type = ExcelColumnType.Int64Column },
                new() { Names = ["v"], Type = type, Nullable = nullable },
            ];
            bool text = type is ExcelColumnType.StringColumn or ExcelColumnType.StringDictionaryColumn;
            if (!nullable && !text)
            {
                InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                    TypedTableParser.Parse(workbook.FirstSheet, specs, 1, dop, allowDictionary: true, 0, out _, out _));
                Assert.Equal("column 1 (\"v\") has a value that failed to convert and is not nullable.", error.Message);
                return;
            }
            using ColumnarTable? table = TypedTableParser.Parse(workbook.FirstSheet, specs, 1, dop, allowDictionary: true, 0, out _, out string? specError);
            Assert.Null(specError);
            Assert.NotNull(table);
            Assert.Equal(4, table.RowCount);
            Assert.Equal(["v:1", "v:2", "v:3", "v:4"], Render(table.Columns[0], ExcelColumnType.Int64Column));
            string valid = Expected(type);
            string[] expected = nullable
                ? [$"v:{valid}", $"n:{Raw(type)}", $"v:{valid}", $"n:{Raw(type)}"]
                : [$"v:{valid}", "v:", $"v:{valid}", "v:"];
            Assert.Equal(expected, Render(table.Columns[1], type));
            Assert.Equal(nullable, table.Columns[1].Validity != IntPtr.Zero);
        }

        [Fact]
        public void Nullable_Columns_Keep_Appending_After_A_Blank()
        {
            using IExcelWorkbook workbook = Excel.FromCsv(Encoding.UTF8.GetBytes("a,b,c\n,,\n1,2.5,x\n,,\n"));
            TypedColumnSpec[] specs =
            [
                new() { Names = ["a"], Type = ExcelColumnType.Int64Column, Nullable = true },
                new() { Names = ["b"], Type = ExcelColumnType.Float64Column, Nullable = true },
                new() { Names = ["c"], Type = ExcelColumnType.StringColumn, Nullable = true },
            ];
            using ColumnarTable table = TypedTableParser.Parse(workbook.FirstSheet, specs, 1, 1, allowDictionary: false, 0, out _, out _)!;
            Assert.Equal(3, table.RowCount);
            Assert.Equal(["n:0", "v:1", "n:0"], Render(table.Columns[0], ExcelColumnType.Int64Column));
            Assert.Equal(["n:0", "v:2.5", "n:0"], Render(table.Columns[1], ExcelColumnType.Float64Column));
            Assert.Equal(["n:", "v:x", "n:"], Render(table.Columns[2], ExcelColumnType.StringColumn));
        }

        [Fact]
        public void Csv_Model_Nullable_Properties_Are_Null_For_Blank_And_Missing_Fields()
        {
            using CsvReader reader = Excel.FromCsv(Encoding.UTF8.GetBytes(
                "Qty,Rate,When,Day,Amount,Flag\n1,2.5,2024-03-05T12:00:00,2024-03-05,9.99,true\n2,,,,,\n,4.5,,,10,\n5,1.5\n"));
            List<BlankRow> rows = [.. ExcelParser.FromAttributes<BlankRow>().Parse(reader.FirstSheet)];
            AssertModel(rows);
        }

        [Fact]
        public async Task Xlsx_Model_Nullable_Properties_Are_Null_For_Blank_And_Missing_Cells()
        {
            await using MemoryStream stream = await TypedWorkbook.BuildAsync(
                ["Qty", "Rate", "When", "Day", "Amount", "Flag"],
                [1, 2.5, Sample, Sample, 9.99m, true],
                [2, null, null, null, null, null],
                [new Gap(), 4.5, new Gap(), new Gap(), 10m],
                [5, 1.5]);
            await using var reader = Excel.FromXlsx(stream);
            List<BlankRow> rows = [.. ExcelParser.FromAttributes<BlankRow>().Parse(reader.FirstSheet)];
            AssertModel(rows);
        }

        private static void AssertModel(List<BlankRow> rows)
        {
            Assert.Equal(4, rows.Count);
            Assert.Equal((1, 2.5, (DateTime?)Sample, (DateOnly?)DateOnly.FromDateTime(Sample), (decimal?)9.99m, (bool?)true),
                (rows[0].Qty, rows[0].Rate, rows[0].When, rows[0].Day, rows[0].Amount, rows[0].Flag));
            Assert.Equal((2, null, null, null, null, null), (rows[1].Qty, rows[1].Rate, rows[1].When, rows[1].Day, rows[1].Amount, rows[1].Flag));
            Assert.Equal((null, 4.5, null, null, (decimal?)10m, null), (rows[2].Qty, rows[2].Rate, rows[2].When, rows[2].Day, rows[2].Amount, rows[2].Flag));
            Assert.Equal((5, 1.5, null, null, null, null), (rows[3].Qty, rows[3].Rate, rows[3].When, rows[3].Day, rows[3].Amount, rows[3].Flag));
        }

        [Fact]
        public void Empty_Cell_Reads_Fail_For_Dates_And_Parsable_Primitives()
        {
            AssertBlankFails(new Cell(CellType.Empty, default));
            AssertBlankFails(new Cell(CellType.ExcelString, default));
        }

        private static void AssertBlankFails(Cell cell)
        {
            Assert.False(cell.TryGetDateTime(isDate1904: false, out DateTime date));
            Assert.Equal(default, date);
            Assert.False(cell.TryParse(null, out double d));
            Assert.False(cell.TryParse(null, out float f));
            Assert.False(cell.TryParse(null, out Half h));
            Assert.False(cell.TryParse(null, out decimal m));
            Assert.False(cell.TryParse(null, out int i));
            Assert.False(cell.TryParse(null, out long l));
            Assert.False(cell.TryParse(null, out byte b));
            Assert.False(cell.TryParse(null, out Guid guid));
            Assert.Equal((0d, 0f, Half.Zero, 0m, 0, 0L, (byte)0, Guid.Empty), (d, f, h, m, i, l, b, guid));
        }

        [Fact]
        public void Empty_Cell_Read_Still_Reaches_A_Type_That_Accepts_It()
        {
            Cell cell = new(CellType.Empty, default);
            Assert.True(cell.TryParse(null, out AcceptsEmpty _));
        }

        [Fact]
        public void Number_Cells_With_Empty_Text_Still_Read()
        {
            Cell cell = new(CellType.Number, default, 45000, hasNumber: true, styleIndex: 0);
            Assert.True(cell.TryGetDateTime(isDate1904: false, out DateTime date));
            Assert.Equal(DateTime.FromOADate(45000), date);
            Assert.True(cell.TryParse(null, out int value));
            Assert.Equal(45000, value);
        }

        private static CsvReader OpenCsv(ExcelColumnType type)
        {
            string sample = type switch
            {
                ExcelColumnType.Int64Column => "7",
                ExcelColumnType.Float64Column => "1.5",
                ExcelColumnType.BoolColumn => "true",
                ExcelColumnType.DateColumn => "2024-03-05",
                ExcelColumnType.TimeColumn => "12:00:00",
                ExcelColumnType.TimestampColumn => "2024-03-05T12:00:00",
                _ => "abc",
            };
            return Excel.FromCsv(Encoding.UTF8.GetBytes($"k,v\n1,{sample}\n2,\n3,{sample}\n4\n"));
        }

        private static XlsxWorkbook OpenXlsx(ExcelColumnType type)
        {
            object sample = type switch
            {
                ExcelColumnType.Int64Column => 7,
                ExcelColumnType.Float64Column => 1.5,
                ExcelColumnType.BoolColumn => true,
                ExcelColumnType.DateColumn or ExcelColumnType.TimeColumn or ExcelColumnType.TimestampColumn => Sample,
                _ => "abc",
            };
            MemoryStream stream = TypedWorkbook.BuildAsync(["k", "v"], [1, sample], [2, null], [3, sample], [4]).GetAwaiter().GetResult();
            return Excel.FromXlsx(stream, leaveOpen: false);
        }

        private static string Expected(ExcelColumnType type)
        {
            return type switch
            {
                ExcelColumnType.Int64Column => "7",
                ExcelColumnType.Float64Column => "1.5",
                ExcelColumnType.BoolColumn => "1",
                ExcelColumnType.DateColumn => (DateOnly.FromDateTime(Sample).DayNumber - DateOnly.FromDateTime(DateTime.UnixEpoch).DayNumber).ToString(CultureInfo.InvariantCulture),
                ExcelColumnType.TimeColumn => (Sample.TimeOfDay.Ticks / 10).ToString(CultureInfo.InvariantCulture),
                ExcelColumnType.TimestampColumn => ((Sample - DateTime.UnixEpoch).Ticks / 10).ToString(CultureInfo.InvariantCulture),
                _ => "abc",
            };
        }

        private static string Raw(ExcelColumnType type)
        {
            return type switch
            {
                ExcelColumnType.StringColumn => "",
                ExcelColumnType.StringDictionaryColumn => "abc",
                _ => "0",
            };
        }

        private static string[] Render(BuiltColumn column, ExcelColumnType type)
        {
            int count = (int)column.Length;
            byte[] validity = new byte[(count + 7) / 8];
            if (column.Validity != IntPtr.Zero)
            {
                Marshal.Copy(column.Validity, validity, 0, validity.Length);
            }
            string[] raw = RawValues(column, type, count);
            string[] rendered = new string[count];
            for (int i = 0; i < count; i++)
            {
                bool valid = column.Validity == IntPtr.Zero || (validity[i >> 3] & (1 << (i & 7))) != 0;
                rendered[i] = (valid ? "v:" : "n:") + raw[i];
            }
            return rendered;
        }

        private static string[] RawValues(BuiltColumn column, ExcelColumnType type, int count)
        {
            switch (type)
            {
                case ExcelColumnType.StringColumn:
                    return Strings(column.Values, column.Data, (int)column.DataLen, count);
                case ExcelColumnType.StringDictionaryColumn:
                    int[] indices = new int[count];
                    Marshal.Copy(column.Values, indices, 0, count);
                    string[] entries = Strings(column.DictionaryOffsets, column.Data, (int)column.DataLen, (int)column.DictionaryLength);
                    return [.. indices.Select(i => entries[i])];
                case ExcelColumnType.Float64Column:
                    double[] doubles = new double[count];
                    Marshal.Copy(column.Values, doubles, 0, count);
                    return [.. doubles.Select(d => d.ToString(CultureInfo.InvariantCulture))];
                case ExcelColumnType.BoolColumn:
                    byte[] bytes = new byte[count];
                    Marshal.Copy(column.Values, bytes, 0, count);
                    return [.. bytes.Select(b => b.ToString(CultureInfo.InvariantCulture))];
                case ExcelColumnType.DateColumn:
                    int[] ints = new int[count];
                    Marshal.Copy(column.Values, ints, 0, count);
                    return [.. ints.Select(i => i.ToString(CultureInfo.InvariantCulture))];
                default:
                    long[] longs = new long[count];
                    Marshal.Copy(column.Values, longs, 0, count);
                    return [.. longs.Select(l => l.ToString(CultureInfo.InvariantCulture))];
            }
        }

        private static string[] Strings(IntPtr offsetsBlock, IntPtr dataBlock, int dataLength, int count)
        {
            int[] offsets = new int[count + 1];
            Marshal.Copy(offsetsBlock, offsets, 0, offsets.Length);
            byte[] data = new byte[dataLength];
            Marshal.Copy(dataBlock, data, 0, dataLength);
            string[] strings = new string[count];
            for (int i = 0; i < count; i++)
            {
                strings[i] = Encoding.UTF8.GetString(data, offsets[i], offsets[i + 1] - offsets[i]);
            }
            return strings;
        }
    }
}
