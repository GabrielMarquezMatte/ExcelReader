using System.Globalization;
using System.Text;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Reader
{
    /// <summary>
    /// A cell borrowed from the native reader, valid only while the callback that received its row runs. A cell
    /// carries text, a number, or both.
    /// </summary>
    public readonly unsafe ref struct ExcelCellView
    {
        private readonly XlCellView* _cell;

        internal ExcelCellView(XlCellView* cell)
        {
            _cell = cell;
        }

        /// <summary>The zero-based column of the cell.</summary>
        public int Column
        {
            get
            {
                return _cell->Column;
            }
        }

        /// <summary>The kind of value the cell holds.</summary>
        public ExcelCellType Type
        {
            get
            {
                return (ExcelCellType)_cell->Type;
            }
        }

        /// <summary>
        /// The cell's number: the native double when the cell carries one, otherwise its text parsed with the invariant
        /// culture; null when neither applies. A <see cref="ExcelCellType.Date"/> cell yields its Excel serial number.
        /// </summary>
        public double? Number
        {
            get
            {
                if (_cell->HasNumber != 0)
                {
                    return _cell->Number;
                }

                if (_cell->ValueLen == 0)
                {
                    return null;
                }

                string text = Encoding.UTF8.GetString(_cell->Value, _cell->ValueLen);
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
                {
                    return null;
                }

                return number;
            }
        }

        /// <summary>
        /// Decodes the cell text into a new string, or formats the number in round-trip form when the cell carries only
        /// a number; null when the cell is empty.
        /// </summary>
        /// <returns>The text, or null.</returns>
        public string? GetText()
        {
            if (_cell->ValueLen > 0)
            {
                return Encoding.UTF8.GetString(_cell->Value, _cell->ValueLen);
            }

            if (_cell->HasNumber == 0)
            {
                return null;
            }

            return _cell->Number.ToString("R", CultureInfo.InvariantCulture);
        }
    }
}
