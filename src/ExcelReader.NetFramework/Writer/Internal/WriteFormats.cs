using ExcelReader.NetFramework.Interop;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Writer.Internal
{
    internal static class WriteFormats
    {
        private const int MaxAscii = 0x7F;

        internal static ExcelFormat FromPath(string path, ExcelFormat? explicitFormat)
        {
            if (explicitFormat is ExcelFormat format)
            {
                return Require(format);
            }

            string extension = Path.GetExtension(path);
            if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelFormat.Xlsx;
            }

            if (string.Equals(extension, ".xlsb", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelFormat.Xlsb;
            }

            if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelFormat.Xls;
            }

            if (string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelFormat.Csv;
            }

            throw new ArgumentException($"Cannot infer the format from the extension '{extension}'; set ExcelWriteOptions.Format.", nameof(path));
        }

        internal static ExcelFormat Require(ExcelFormat format)
        {
            if (format == ExcelFormat.Auto || !Enum.IsDefined(typeof(ExcelFormat), format))
            {
                throw new ArgumentException("The format must be Xls, Xlsx, Xlsb or Csv.", nameof(format));
            }

            return format;
        }

        internal static string? PasswordFor(ExcelFormat format, ExcelWriteOptions? options)
        {
            string? password = options?.Password;
            if (password is not null && password.Length == 0)
            {
                throw EmptyPassword(nameof(options.Password));
            }

            if (password is not null && format != ExcelFormat.Xlsx && format != ExcelFormat.Xlsb)
            {
                throw new NotSupportedException("A password can encrypt XLSX and XLSB output only.");
            }

            return password;
        }

        internal static unsafe XlWriteOptions ToNative(ExcelWriteOptions? options)
        {
            XlWriteOptions native = new XlWriteOptions { StructSize = sizeof(XlWriteOptions) };

            if (options is null)
            {
                return native;
            }

            RequireNotNegative(options.DegreeOfParallelism, nameof(options.DegreeOfParallelism));
            RequireNotNegative(options.BatchRows, nameof(options.BatchRows));
            native.CsvDelimiter = AsciiOrDefault(options.CsvDelimiter, nameof(options.CsvDelimiter));
            native.CsvQuote = AsciiOrDefault(options.CsvQuote, nameof(options.CsvQuote));
            native.Date1904 = Flag(options.Date1904);
            native.UseSharedStrings = Flag(options.UseSharedStrings);
            native.DegreeOfParallelism = options.DegreeOfParallelism;
            native.BatchRows = options.BatchRows;
            return native;
        }

        private static ArgumentException EmptyPassword(string name)
        {
            return new ArgumentException("The password cannot be empty.", name);
        }

        private static void RequireNotNegative(int value, string name)
        {
            if (value < 0)
            {
                throw new ArgumentOutOfRangeException(name, value, "The value cannot be negative.");
            }
        }

        private static int Flag(bool? value)
        {
            if (value is null)
            {
                return Xl.XL_OPT_DEFAULT;
            }

            return value.Value ? Xl.XL_OPT_TRUE : Xl.XL_OPT_FALSE;
        }

        private static int AsciiOrDefault(char? value, string name)
        {
            if (value is null)
            {
                return 0;
            }

            if (value > MaxAscii)
            {
                throw new ArgumentOutOfRangeException(name, value, "The native library takes a single-byte character.");
            }

            return value.Value;
        }
    }
}
