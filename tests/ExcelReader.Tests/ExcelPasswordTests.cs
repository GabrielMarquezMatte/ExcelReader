using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class ExcelPasswordTests
    {
        [Fact]
        public void Should_Redact_When_Converted_To_String()
        {
            var password = new ExcelPassword("hunter2");
            Assert.DoesNotContain("hunter2", password.ToString(), StringComparison.Ordinal);
        }

        // ExcelReaderOptions is a record: its synthesized ToString prints every property, so a bare
        // string Password would leak into any structured log of the options object.
        [Fact]
        public void Should_Not_Leak_Password_When_Options_Are_Formatted()
        {
            var options = ExcelReaderOptions.Default with { Password = "hunter2" };
            Assert.DoesNotContain("hunter2", options.ToString(), StringComparison.Ordinal);
        }

        [Fact]
        public void Should_Convert_Implicitly_When_Assigned_From_String()
        {
            ExcelReaderOptions options = ExcelReaderOptions.Default with { Password = "hunter2" };
            Assert.NotNull(options.Password);
        }

        [Fact]
        public void Should_Accept_Span_When_Caller_Avoids_A_String()
        {
            var password = new ExcelPassword("hunter2".AsSpan());
            Assert.DoesNotContain("hunter2", password.ToString(), StringComparison.Ordinal);
        }

        [Fact]
        public void Should_Default_To_Null_When_Not_Set()
        {
            Assert.Null(ExcelReaderOptions.Default.Password);
            Assert.False(ExcelReaderOptions.Default.VerifyEncryptedIntegrity);
            Assert.Equal(100_000, ExcelReaderOptions.Default.MaxPasswordSpinCount);
        }
    }
}
