using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader
{
    public sealed class SchemaInferenceTests
    {
        [Fact]
        public void Should_MatchTheNativeTypeCodes_When_CastingExcelColumnType()
        {
            Assert.Equal(0, (int)ExcelColumnType.StringColumn);
            Assert.Equal(1, (int)ExcelColumnType.Int64Column);
            Assert.Equal(2, (int)ExcelColumnType.Float64Column);
            Assert.Equal(3, (int)ExcelColumnType.BoolColumn);
            Assert.Equal(4, (int)ExcelColumnType.DateColumn);
            Assert.Equal(5, (int)ExcelColumnType.TimeColumn);
            Assert.Equal(6, (int)ExcelColumnType.TimestampColumn);
        }

        [Fact]
        public void Should_CompareByValue_When_TwoSchemasDescribeTheSameColumn()
        {
            ExcelColumnSchema a = new() { Index = 2, Name = "Total", Type = ExcelColumnType.Float64Column, IsNullable = true };
            ExcelColumnSchema b = new() { Index = 2, Name = "Total", Type = ExcelColumnType.Float64Column, IsNullable = true };
            ExcelColumnSchema c = a with { IsNullable = false };

            Assert.Equal(a, b);
            Assert.Equal(a.GetHashCode(), b.GetHashCode());
            Assert.NotEqual(a, c);
        }
        private static ExcelColumnSchema[] InferFromCsv(string csv, int headerRow = 1, int sampleSize = 100, bool parseText = false)
        {
            using CsvReader reader = Excel.FromCsv(System.Text.Encoding.UTF8.GetBytes(csv));
            using var rows = reader.GetEnumerator();
            return SchemaInference.Infer(rows, isDate1904: false, headerRow, sampleSize, parseText);
        }

        private static ExcelColumnSchema[] InferFromXlsx(string sheetRows, int headerRow = 1, int sampleSize = 100, bool parseText = false)
        {
            using MemoryStream ms = WorkbookBuilder.Build(sheetRows);
            using XlsxReader reader = Excel.FromXlsx(ms.ToArray());
            using var rows = reader.GetEnumerator();
            return SchemaInference.Infer(rows, isDate1904: false, headerRow, sampleSize, parseText);
        }

        private static ExcelColumnType[] TypesOf(ExcelColumnSchema[] schema)
        {
            return Array.ConvertAll(schema, column => column.Type);
        }

        [Fact]
        public void ParseText_Should_Infer_Integers_And_Decimals_But_Not_Codes()
        {
            ExcelColumnSchema[] schema = InferFromCsv("Qty,Ratio,Code,Signed\n1,0.5,00123,-0\n2,1,00456,+5\n", parseText: true);

            Assert.Equal(
                [ExcelColumnType.Int64Column, ExcelColumnType.Float64Column, ExcelColumnType.StringColumn, ExcelColumnType.Int64Column],
                TypesOf(schema));
        }

        [Fact]
        public void ParseText_Should_Infer_Booleans_Without_Swallowing_Binary_Integers()
        {
            ExcelColumnSchema[] schema = InferFromCsv("A,B,C,D\ntrue,0,true,true\nFALSE,1,1,2\n", parseText: true);

            Assert.Equal(
                [ExcelColumnType.BoolColumn, ExcelColumnType.Int64Column, ExcelColumnType.BoolColumn, ExcelColumnType.StringColumn],
                TypesOf(schema));
        }

        [Fact]
        public void ParseText_Should_Infer_Iso_Dates_And_Timestamps_But_Not_Loose_Dates()
        {
            ExcelColumnSchema[] schema = InferFromCsv(
                "Day,Stamp,Precise,Slash,Clock\n"
                + "2024-01-02,2024-01-02,2024-01-02T10:30:00.5,1/2,10:00\n"
                + "2024-02-03,2024-02-03 10:30:00,2024-02-03T11:00:00,3/4,11:00\n",
                parseText: true);

            Assert.Equal(
                [ExcelColumnType.DateColumn, ExcelColumnType.TimestampColumn, ExcelColumnType.TimestampColumn,
                    ExcelColumnType.StringColumn, ExcelColumnType.StringColumn],
                TypesOf(schema));
        }

        [Fact]
        public void ParseText_Should_Keep_Text_That_Only_Loosely_Looks_Numeric()
        {
            ExcelColumnSchema[] schema = InferFromCsv(
                "Padded,Huge,Word,Nan\n\" 42\",99999999999999999999,1,NaN\n\"7 \",1,text,1.5\n",
                parseText: true);

            Assert.All(schema, column => Assert.Equal(ExcelColumnType.StringColumn, column.Type));
        }

        [Fact]
        public void ParseText_Should_Mark_Nullable_Columns_And_Type_Them_From_Their_Values()
        {
            ExcelColumnSchema[] schema = InferFromCsv("Qty,Name\n1,a\n,b\n3,c\n", parseText: true);

            Assert.Equal(ExcelColumnType.Int64Column, schema[0].Type);
            Assert.True(schema[0].IsNullable);
        }

        [Fact]
        public void ParseText_Should_Keep_An_Empty_Column_As_Nullable_String()
        {
            ExcelColumnSchema[] schema = InferFromCsv("Qty,Blank\n1,\n2,\n", parseText: true);

            Assert.Equal(ExcelColumnType.StringColumn, schema[1].Type);
            Assert.True(schema[1].IsNullable);
        }

        [Fact]
        public void ParseText_Off_Should_Leave_Every_Csv_Column_A_String()
        {
            ExcelColumnSchema[] schema = InferFromCsv("Qty,Ratio,Flag,Day\n1,0.5,true,2024-01-02\n2,1.5,false,2024-01-03\n");

            Assert.All(schema, column => Assert.Equal(ExcelColumnType.StringColumn, column.Type));
        }

        [Fact]
        public void ParseText_Should_Merge_Tagged_And_Text_Cells_In_Xlsx_Only_When_Parsing_Text()
        {
            const string Rows =
                """
                <row r="1">
                    <c r="A1" t="inlineStr"><is><t>Qty</t></is></c>
                    <c r="B1" t="inlineStr"><is><t>Flag</t></is></c>
                    <c r="C1" t="inlineStr"><is><t>Odd</t></is></c>
                </row>
                <row r="2">
                    <c r="A2"><v>1</v></c>
                    <c r="B2" t="b"><v>1</v></c>
                    <c r="C2" t="b"><v>1</v></c>
                </row>
                <row r="3">
                    <c r="A3" t="inlineStr"><is><t>2</t></is></c>
                    <c r="B3"><v>0</v></c>
                    <c r="C3"><v>2</v></c>
                </row>
                """;

            Assert.Equal(
                [ExcelColumnType.Int64Column, ExcelColumnType.BoolColumn, ExcelColumnType.StringColumn],
                TypesOf(InferFromXlsx(Rows, parseText: true)));
            Assert.Equal(
                [ExcelColumnType.StringColumn, ExcelColumnType.StringColumn, ExcelColumnType.StringColumn],
                TypesOf(InferFromXlsx(Rows)));
        }

        [Fact]
        public void Should_GuessOneTypePerColumn_When_EverySampledCellAgrees()
        {
            ExcelColumnSchema[] schema = InferFromXlsx(
                """
                <row r="1">
                    <c r="A1" t="inlineStr"><is><t>Id</t></is></c>
                    <c r="B1" t="inlineStr"><is><t>Name</t></is></c>
                    <c r="C1" t="inlineStr"><is><t>Ratio</t></is></c>
                </row>
                <row r="2">
                    <c r="A2"><v>1</v></c>
                    <c r="B2" t="inlineStr"><is><t>alice</t></is></c>
                    <c r="C2"><v>0.5</v></c>
                </row>
                <row r="3">
                    <c r="A3"><v>2</v></c>
                    <c r="B3" t="inlineStr"><is><t>bob</t></is></c>
                    <c r="C3"><v>1.25</v></c>
                </row>
                """);

            Assert.Equal(3, schema.Length);
            Assert.Equal(new ExcelColumnSchema { Index = 0, Name = "Id", Type = ExcelColumnType.Int64Column, IsNullable = false }, schema[0]);
            Assert.Equal(new ExcelColumnSchema { Index = 1, Name = "Name", Type = ExcelColumnType.StringColumn, IsNullable = false }, schema[1]);
            Assert.Equal(new ExcelColumnSchema { Index = 2, Name = "Ratio", Type = ExcelColumnType.Float64Column, IsNullable = false }, schema[2]);
        }

        [Fact]
        public void Should_FallBackToString_When_SampledCellsMixKinds()
        {
            ExcelColumnSchema[] schema = InferFromCsv("Mixed\n1\ntext\n");

            Assert.Equal(ExcelColumnType.StringColumn, schema[0].Type);
        }

        [Fact]
        public void Should_MarkNullable_When_ARowLeavesTheColumnEmpty()
        {
            ExcelColumnSchema[] schema = InferFromCsv("Id,Note\n1,here\n2,\n");

            Assert.False(schema[0].IsNullable);
            Assert.True(schema[1].IsNullable);
        }

        [Fact]
        public void Should_LeaveNameNull_When_HeaderRowIsZero()
        {
            ExcelColumnSchema[] schema = InferFromCsv("1,2\n3,4\n", headerRow: 0);

            Assert.All(schema, column => Assert.Null(column.Name));
            Assert.Equal(0, schema[0].Index);
            Assert.Equal(1, schema[1].Index);
        }

        [Fact]
        public void Should_LeaveNameNull_When_AHeaderCellIsBlank()
        {
            ExcelColumnSchema[] schema = InferFromCsv("Id,,Tail\n1,2,3\n");

            Assert.Equal("Id", schema[0].Name);
            Assert.Null(schema[1].Name);
            Assert.Equal("Tail", schema[2].Name);
        }

        [Fact]
        public void Should_Throw_When_TheSheetHasFewerRowsThanTheHeaderRow()
        {
            ArgumentException error = Assert.Throws<ArgumentException>(
                () => InferFromCsv("Id\n1\n", headerRow: 99));

            Assert.Contains("99", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_SampleAtMostSampleSizeRows_When_TheSheetIsLonger()
        {
            ExcelColumnSchema[] schema = InferFromXlsx(
                """
                <row r="1"><c r="A1" t="inlineStr"><is><t>Id</t></is></c></row>
                <row r="2"><c r="A2"><v>1</v></c></row>
                <row r="3"><c r="A3" t="inlineStr"><is><t>text</t></is></c></row>
                """,
                sampleSize: 1);

            Assert.Equal(ExcelColumnType.Int64Column, schema[0].Type);
        }

        [Fact]
        public void Should_InferAcrossEveryFormat_When_CalledThroughTheExcelFacade()
        {
            using IExcelRowReader reader = Excel.Open(Path.Combine("data", "RealExcel.xlsb"));
            ExcelColumnSchema[] schema = Excel.InferSchema(reader);

            Assert.NotEmpty(schema);
            Assert.All(schema, column => Assert.True(column.Index >= 0));
            for (int i = 0; i < schema.Length; i++)
            {
                Assert.Equal(i, schema[i].Index);
            }
        }

        [Fact]
        public void Should_NotDisturbTheReader_When_InferSchemaRunsBeforeEnumeration()
        {
            using IExcelRowReader reader = Excel.FromCsv(System.Text.Encoding.UTF8.GetBytes("Id\n1\n2\n"));
            _ = Excel.InferSchema(reader);

            using IExcelRowEnumerator rows = reader.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal("Id", rows.Current[0].GetString());
        }

        [Fact]
        public void Should_Throw_When_ReaderIsNull()
        {
            Assert.Throws<ArgumentNullException>(() => Excel.InferSchema(null!));
        }
    }
}
