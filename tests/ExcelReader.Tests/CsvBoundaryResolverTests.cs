using ExcelReader.Core.Parser.Internal;

namespace ExcelReader.Tests
{
    public class CsvBoundaryResolverTests
    {
        private const byte Quote = (byte)'"';

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
            //          0123456789
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

            // Outside: the \n at index 1 ends a record.
            Assert.Equal(2, outside);
            // Inside: index 1's \n is inside quotes; the quote at index 3 closes the field, so the
            // \n at index 4 is the first real terminator.
            Assert.Equal(5, inside);
        }

        [Fact]
        public void ADoubledQuoteLeavesTheParityUnchanged()
        {
            // Starting inside: "" is an escaped quote, so we stay inside and the \n is literal.
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
            // Whether or not a \n follows, the record would start at or past the window's end, so
            // there is nothing in this window to parse.
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
    }
}
