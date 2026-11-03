namespace ExcelReader.NetFramework.Tests
{
    public sealed class Row
    {
        [ExcelColumn("Coluna1")]
        public string? Coluna1 { get; set; }

        [ExcelColumn("Coluna2")]
        public DateTime Coluna2 { get; set; }

        [ExcelColumn("Coluna3")]
        public int Coluna3 { get; set; }

        [ExcelColumn("Coluna16")]
        public double Coluna16 { get; set; }
    }
}
