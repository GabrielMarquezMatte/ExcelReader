using System.Text;
using ExcelReader.Core.Parser.ParallelCsv;

namespace ExcelReader.Tests.Parser.ParallelCsv
{
    public class CsvBoundaryResolverTests
    {
        private const byte Quote = (byte)'"';
        private const byte Delimiter = (byte)',';

        [Fact]
        public void FindsTheByteAfterAPlainNewline()
        {
            ReadOnlySpan<byte> window = "abc\ndef"u8;

            int start = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Outside);

            Assert.Equal(4, start);
        }

        [Fact]
        public void SkipsBothBytesOfACrLfTerminator()
        {
            ReadOnlySpan<byte> window = "abc\r\ndef"u8;

            int start = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Outside);

            Assert.Equal(5, start);
        }

        [Fact]
        public void TreatsALoneCarriageReturnAsATerminator()
        {
            ReadOnlySpan<byte> window = "abc\rdef"u8;

            int start = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Outside);

            Assert.Equal(4, start);
        }

        [Fact]
        public void IgnoresANewlineInsideAQuotedFieldWhenStartingOutside()
        {
            ReadOnlySpan<byte> window = "\"a\nb\"\nc"u8;

            int start = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Outside);

            Assert.Equal(6, start);
        }

        [Fact]
        public void TheInsideHypothesisTakesTheNewlineTheOutsideHypothesisSkips()
        {
            ReadOnlySpan<byte> window = "a\nb\"\nc"u8;

            int outside = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Outside);
            int inside = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Inside);

            Assert.Equal(2, outside);
            Assert.Equal(5, inside);
        }

        [Fact]
        public void ADoubledQuoteLeavesTheParityUnchanged()
        {
            ReadOnlySpan<byte> window = "a\"\"b\nc\"\nd"u8;

            int inside = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Inside);

            Assert.Equal(8, inside);
        }

        [Fact]
        public void ReturnsMinusOneWhenTheWindowHoldsNoBoundary()
        {
            ReadOnlySpan<byte> window = "abcdef"u8;

            int start = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Outside);

            Assert.Equal(-1, start);
        }

        [Fact]
        public void ATrailingTerminatorLeavesNoRecordInsideTheWindow()
        {
            Assert.Equal(-1, CsvBoundaryResolver.FindRecordStart("abc\r"u8, Quote, CsvQuoteParity.Outside));
            Assert.Equal(-1, CsvBoundaryResolver.FindRecordStart("abc\n"u8, Quote, CsvQuoteParity.Outside));
            Assert.Equal(-1, CsvBoundaryResolver.FindRecordStart("abc\r\n"u8, Quote, CsvQuoteParity.Outside));
        }

        [Fact]
        public void FindsABoundaryThatLandsOnTheVeryLastByte()
        {
            ReadOnlySpan<byte> window = "ab\nc"u8;

            int start = CsvBoundaryResolver.FindRecordStart(window, Quote, CsvQuoteParity.Outside);

            Assert.Equal(3, start);
        }

        private static CsvQuoteParity? Infer(string window, char before = 'x', bool endsAtEof = false, byte delimiter = Delimiter)
        {
            return InferAfter(before.ToString(), window, endsAtEof, delimiter);
        }

        private static CsvQuoteParity? InferAfter(string behind, string window, bool endsAtEof = false, byte delimiter = Delimiter)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(behind + window);
            return CsvBoundaryResolver.InferParity(bytes, Encoding.UTF8.GetByteCount(behind), startsAtSource: false, endsAtEof, Quote, delimiter);
        }

        [Theory]
        [InlineData("th\nree\",4\n")]
        [InlineData("ree\"\n")]
        [InlineData("ree\"\r\n")]
        [InlineData("a\"\"b\",4\n")]
        [InlineData("ab\"\"\",4\n")]
        public void ARunThatClosesAFieldMeansTheWindowStartsInside(string window)
        {
            Assert.Equal(CsvQuoteParity.Inside, Infer(window));
        }

        [Theory]
        [InlineData("4,\"three\",5\n")]
        [InlineData(",4\n\"x\ny\",5\n")]
        [InlineData("4,\"\",5\n")]
        [InlineData("4,\"\"\"x\",5\n")]
        [InlineData("plain,4\nplain,5\n")]
        [InlineData("")]
        [InlineData("ab\"")]
        public void ARunThatOpensAFieldOrNoRunMeansOutside(string window)
        {
            Assert.Equal(CsvQuoteParity.Outside, Infer(window));
        }

        [Theory]
        [InlineData("a,b\"c,4\n")]
        [InlineData("a\"x\",4\n")]
        public void ARunThatBreaksTheQuotingRulesIsGuessedOutside(string window)
        {
            Assert.Equal(CsvQuoteParity.Outside, Infer(window));
        }

        [Theory]
        [InlineData("x,\",b\",4\n")]
        [InlineData("x\n\",b\",4\n")]
        [InlineData("x,\"\n\",4\n")]
        [InlineData("1,\"\"\",x\",2\n")]
        public void ARunThatCouldOpenOrCloseIsUnknown(string window)
        {
            Assert.Null(Infer(window));
        }

        [Fact]
        public void ARunAtTheWindowStartIsReadAgainstTheByteBeforeIt()
        {
            Assert.Equal(CsvQuoteParity.Outside, Infer("\"x\",4\n", before: ','));
            Assert.Equal(CsvQuoteParity.Inside, Infer("\",4\n", before: 'a'));
            Assert.Null(Infer("\",4\n", before: ','));
            Assert.Null(Infer("\"x\",4\n", before: '"'));
            Assert.Null(InferAfter("", "\"x\",4\n"));
        }

        [Fact]
        public void AStartInsideARunCountsTheQuotesBeforeIt()
        {
            Assert.Equal(CsvQuoteParity.Inside, InferAfter("x,\"", "\",4,\"abc\"\n"));
            Assert.Equal(CsvQuoteParity.Outside, InferAfter("x,\"said \"", "\"hi\"\"\",5\n"));
            Assert.Equal(CsvQuoteParity.Inside, InferAfter("x\n\"", "\"\"quoted\"\" start\",5\n"));
            Assert.Null(InferAfter("\"", "\"x\",4\n"));
            Assert.Null(InferAfter("x,\"", "\"\""));
            Assert.Null(InferAfter("a\"", "\"b,4\n"));
        }

        [Fact]
        public void ARunAtTheEndOfTheSourceCloses()
        {
            Assert.Equal(CsvQuoteParity.Inside, Infer("ree\"", endsAtEof: true));
            Assert.Equal(CsvQuoteParity.Outside, Infer("ree\"", endsAtEof: false));
        }

        [Fact]
        public void HonoursTheDelimiter()
        {
            Assert.Equal(CsvQuoteParity.Inside, Infer("ree\";4\n", delimiter: (byte)';'));
            Assert.Equal(CsvQuoteParity.Outside, Infer("ree\";4\n"));
        }
    }
}
