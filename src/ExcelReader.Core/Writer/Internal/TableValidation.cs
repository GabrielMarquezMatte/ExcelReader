using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class TableValidation
    {
        private const int MaxNameLength = 255;

        private static readonly HashSet<string> BuiltInStyles = BuildStyles();

        internal static string[] Validate(string name, IReadOnlyList<string> columns, ExcelTableOptions options)
        {
            ValidateName(name);
            ArgumentNullException.ThrowIfNull(options);
            ArgumentOutOfRangeException.ThrowIfNegative(options.FirstColumn, nameof(options));
            ValidateStyle(options);
            return ValidateColumns(columns, options.FirstColumn);
        }

        private static void ValidateName(string name)
        {
            ArgumentNullException.ThrowIfNull(name);
            if (name.Length is 0 or > MaxNameLength)
            {
                throw new ArgumentException("A table name must be 1 to 255 characters long.", nameof(name));
            }
            if (!IsNameStart(name[0]) || !HasOnlyNameCharacters(name.AsSpan(1)))
            {
                throw new ArgumentException(
                    $"'{name}' is not a valid table name: it must start with a letter, '_' or '\\' and contain only letters, digits, '.' and '_'.",
                    nameof(name));
            }
            if (IsA1Reference(name) || IsR1C1Reference(name))
            {
                throw new ArgumentException($"'{name}' is not a valid table name: Excel reads it as a cell reference.", nameof(name));
            }
        }

        private static bool IsNameStart(char c)
        {
            return char.IsLetter(c) || c is '_' or '\\';
        }

        private static bool HasOnlyNameCharacters(ReadOnlySpan<char> rest)
        {
            foreach (char c in rest)
            {
                if (!char.IsLetterOrDigit(c) && c is not ('.' or '_'))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool IsA1Reference(string name)
        {
            string upper = name.ToUpperInvariant();
            return Ascii.IsValid(upper) && TablePart.TryParseRange(Encoding.ASCII.GetBytes(upper), out _, out _, out _, out _);
        }

        private static bool IsR1C1Reference(string name)
        {
            string upper = name.ToUpperInvariant();
            int i = 0;
            bool marker = false;
            if (i < upper.Length && upper[i] == 'R')
            {
                i = SkipDigits(upper, i + 1);
                marker = true;
            }
            if (i < upper.Length && upper[i] == 'C')
            {
                i = SkipDigits(upper, i + 1);
                marker = true;
            }
            return marker && i == upper.Length;
        }

        private static int SkipDigits(string text, int i)
        {
            while (i < text.Length && char.IsAsciiDigit(text[i]))
            {
                i++;
            }
            return i;
        }

        private static void ValidateStyle(ExcelTableOptions options)
        {
            string? styleName = options.StyleName;
            if (styleName is null || BuiltInStyles.Contains(styleName))
            {
                return;
            }
            throw new ArgumentException(
                $"'{styleName}' is not a built-in table style (TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11).",
                nameof(options));
        }

        private static string[] ValidateColumns(IReadOnlyList<string> columns, int firstColumn)
        {
            ArgumentNullException.ThrowIfNull(columns);
            if (columns.Count == 0)
            {
                throw new ArgumentException("A table needs at least one column.", nameof(columns));
            }
            if (columns.Count > ExcelLimits.MaxColumns - firstColumn)
            {
                throw new ArgumentException(
                    $"A table starting at column {firstColumn} cannot have {columns.Count} columns: the sheet has {ExcelLimits.MaxColumns}.",
                    nameof(columns));
            }
            string[] names = new string[columns.Count];
            HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < names.Length; i++)
            {
                names[i] = ValidateColumnName(columns, i);
                if (!seen.Add(names[i]))
                {
                    throw new ArgumentException($"The column name '{names[i]}' appears more than once.", nameof(columns));
                }
            }
            return names;
        }

        private static string ValidateColumnName(IReadOnlyList<string> columns, int index)
        {
            string? column = columns[index];
            if (string.IsNullOrEmpty(column) || column.Length > MaxNameLength || column.AsSpan().ContainsAnyInRange('\0', '\u001F'))
            {
                throw new ArgumentException("Table column names must be 1 to 255 characters long, with no control characters.", nameof(columns));
            }
            return column;
        }

        private static HashSet<string> BuildStyles()
        {
            HashSet<string> styles = new(StringComparer.Ordinal);
            AddStyles(styles, "TableStyleLight", 21);
            AddStyles(styles, "TableStyleMedium", 28);
            AddStyles(styles, "TableStyleDark", 11);
            return styles;
        }

        private static void AddStyles(HashSet<string> styles, string prefix, int count)
        {
            for (int i = 1; i <= count; i++)
            {
                styles.Add(string.Create(CultureInfo.InvariantCulture, $"{prefix}{i}"));
            }
        }
    }
}
