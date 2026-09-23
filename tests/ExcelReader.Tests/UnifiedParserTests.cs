using System.Collections.Generic;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public partial class UnifiedParserTests
    {
        private static readonly byte[] Csv = "Name,Id,Value\nAda,36,1.5\nGrace,45,2.5\n"u8.ToArray();

        [ExcelSerializable]
        public ref partial struct MappedSale
        {
            [ExcelColumn("Name")]
            public ReadOnlySpan<byte> Name { get; set; }

            [ExcelColumn("Id")]
            public int Id { get; set; }
        }

        private ref struct FluentSale
        {
            public ReadOnlySpan<byte> Name { get; set; }
            public int Id { get; set; }
        }

        private sealed class SaleClass
        {
            public string? Name { get; set; }
            public int Id { get; set; }
            public double Value { get; set; }
        }

        private struct SaleStruct
        {
            public string? Name { get; set; }
            public int Id { get; set; }
            public double Value { get; set; }
        }

        private readonly ref struct SaleRef
        {
            public ReadOnlySpan<byte> Name { get; init; }
            public int Id { get; init; }
            public double Value { get; init; }
        }

        [Fact]
        public void ClassModelStillParsesSynchronously()
        {
            using CsvReader reader = Excel.FromCsv(Csv);
            var names = new List<string?>();
            int ids = 0;
            foreach (SaleClass sale in ExcelParser.FromAttributes<SaleClass>().Parse(reader))
            {
                names.Add(sale.Name);
                ids += sale.Id;
            }
            Assert.Equal(["Ada", "Grace"], names);
            Assert.Equal(81, ids);
        }

        [Fact]
        public void StructModelStillParsesSynchronously()
        {
            using CsvReader reader = Excel.FromCsv(Csv);
            double total = 0;
            int rows = 0;
            foreach (SaleStruct sale in ExcelParser.FromAttributes<SaleStruct>().Parse(reader))
            {
                total += sale.Value;
                rows++;
            }
            Assert.Equal(2, rows);
            Assert.Equal(4.0, total);
        }

        [Fact]
        public void RefStructModelParsesThroughTheStandardParser()
        {
            using CsvReader reader = Excel.FromCsv(Csv);
            var names = new List<string>();
            int ids = 0;
            foreach (SaleRef sale in ExcelParser.FromAttributes<SaleRef>().Parse(reader))
            {
                names.Add(Encoding.UTF8.GetString(sale.Name));
                ids += sale.Id;
            }
            Assert.Equal(["Ada", "Grace"], names);
            Assert.Equal(81, ids);
        }

        [Fact]
        public async Task RefStructModelParsesAsynchronously()
        {
            await using CsvReader reader = Excel.FromCsv(Csv);
            var names = new List<string>();
            double total = 0;
            await foreach (SaleRef sale in ExcelParser.FromAttributes<SaleRef>().ParseAsync(reader, TestContext.Current.CancellationToken))
            {
                names.Add(Encoding.UTF8.GetString(sale.Name));
                total += sale.Value;
            }
            Assert.Equal(["Ada", "Grace"], names);
            Assert.Equal(4.0, total);
        }

        [Fact]
        public void RefStructModelFlowsThroughIEnumerableOfT()
        {
            using CsvReader reader = Excel.FromCsv(Csv);
            IEnumerable<SaleRef> sequence = ExcelParser.FromAttributes<SaleRef>().Parse(reader);
            int rows = 0;
            foreach (SaleRef sale in sequence)
            {
                rows++;
                Assert.False(sale.Name.IsEmpty);
            }
            Assert.Equal(2, rows);
        }

        [Fact]
        public async Task RefStructModelFlowsThroughIAsyncEnumerableOfT()
        {
            await using CsvReader reader = Excel.FromCsv(Csv);
            IAsyncEnumerable<SaleRef> sequence = ExcelParser.FromAttributes<SaleRef>().ParseAsync(reader, TestContext.Current.CancellationToken);
            int rows = 0;
            await foreach (SaleRef sale in sequence)
            {
                rows++;
                Assert.False(sale.Name.IsEmpty);
            }
            Assert.Equal(2, rows);
        }

        [Fact]
        public void RefStructModelParsesThroughExcelMappedParser()
        {
            using CsvReader reader = Excel.FromCsv(Csv);
            var names = new List<string>();
            int ids = 0;
            foreach (MappedSale sale in ExcelParser.Generated<MappedSale>().Parse(reader))
            {
                names.Add(Encoding.UTF8.GetString(sale.Name));
                ids += sale.Id;
            }
            Assert.Equal(["Ada", "Grace"], names);
            Assert.Equal(81, ids);
        }

        [Fact]
        public void RefStructModelParsesThroughExcelFluentParser()
        {
            using CsvReader reader = Excel.FromCsv(Csv);
            var parser = ExcelParser.Build<FluentSale>(static builder => builder
                .PropertyRaw(["Name"], static (ref FluentSale m, in Cell c, bool _, IFormatProvider _) =>
                {
                    m.Name = c.Value;
                    return true;
                })
                .PropertyRaw(["Id"], static (ref FluentSale m, in Cell c, bool _, IFormatProvider p) =>
                {
                    bool ok = c.TryParse(p, out int id);
                    m.Id = id;
                    return ok;
                }));

            var names = new List<string>();
            int ids = 0;
            foreach (FluentSale sale in parser.Parse(reader))
            {
                names.Add(Encoding.UTF8.GetString(sale.Name));
                ids += sale.Id;
            }
            Assert.Equal(["Ada", "Grace"], names);
            Assert.Equal(81, ids);
        }

        [Fact]
        public void NonGenericCurrentIsNoLongerSupported()
        {
            using CsvReader reader = Excel.FromCsv(Csv);
            CsvEnumerable<SaleClass>.Enumerator e = ExcelParser.FromAttributes<SaleClass>().Parse(reader).GetEnumerator();
            Assert.True(e.MoveNext());
            System.Collections.IEnumerator legacy = e;
            Assert.Throws<NotSupportedException>(() => legacy.Current);
        }
    }
}
