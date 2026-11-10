using System.Text;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader.Internal
{
    public sealed class Utf8StringsTests
    {
        public static TheoryData<int> Lengths()
        {
            TheoryData<int> lengths = [];
            for (int length = 0; length <= 20; length++)
            {
                lengths.Add(length);
            }
            return lengths;
        }

        [Theory]
        [MemberData(nameof(Lengths))]
        public void AsciiMatchesEncoding(int length)
        {
            byte[] utf8 = new byte[length];
            for (int i = 0; i < length; i++)
            {
                utf8[i] = (byte)(0x20 + (i * 7 % 0x5F));
            }
            Assert.Equal(Encoding.UTF8.GetString(utf8), Utf8Strings.Create(utf8));
        }

        [Theory]
        [MemberData(nameof(Lengths))]
        public void NonAsciiAtAnyPositionMatchesEncoding(int length)
        {
            for (int position = 0; position < length; position++)
            {
                byte[] utf8 = Encoding.UTF8.GetBytes(new string('a', length));
                utf8[position] = 0xFF;
                Assert.Equal(Encoding.UTF8.GetString(utf8), Utf8Strings.Create(utf8));
            }
        }

        [Theory]
        [InlineData("é")]
        [InlineData("aé")]
        [InlineData("São Paulo")]
        [InlineData("abcdefghijklmé")]
        [InlineData("€")]
        public void MultiByteTextMatchesEncoding(string text)
        {
            Assert.Equal(text, Utf8Strings.Create(Encoding.UTF8.GetBytes(text)));
        }

        [Fact]
        public void ValueEndingInsideALargerBufferReadsOnlyItsOwnBytes()
        {
            byte[] buffer = Encoding.UTF8.GetBytes("abcdeéééé");
            Assert.Equal("bcd", Utf8Strings.Create(buffer.AsSpan(1, 3)));
            Assert.Equal("abcde", Utf8Strings.Create(buffer.AsSpan(0, 5)));
        }
    }
}
