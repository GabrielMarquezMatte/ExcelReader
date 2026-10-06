using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Tests.Reader.Csv
{
    public sealed class CsvStructuralScannerTests
    {
        [Fact]
        public void PrefixXor_Matches_The_Software_Ladder()
        {
            ulong[] edges = [0, 1, 1UL << 63, ulong.MaxValue, 0x5555_5555_5555_5555, 0x8000_0000_0000_0001];
            foreach (ulong edge in edges)
            {
                Assert.Equal(CsvStructuralScanner.PrefixXorSoftware(edge), CsvStructuralScanner.PrefixXor(edge));
            }
            ulong bits = 1;
            for (int i = 0; i < 10_000; i++)
            {
                bits = (bits * 6364136223846793005UL) + 1442695040888963407UL;
                Assert.Equal(CsvStructuralScanner.PrefixXorSoftware(bits), CsvStructuralScanner.PrefixXor(bits));
            }
        }
    }
}
