using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;
using static ExcelReader.Core.Reader.Biff12;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsReader
    {
        /// <summary>Enumerates the rows of a single worksheet in an <see cref="XlsReader"/>, supporting both synchronous and asynchronous iteration.</summary>
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : IExcelRowEnumerator
        {
            private readonly XlsReader _reader;
            private readonly CancellationToken _ct;
            private readonly BiffCursor _cursor;
            private readonly CellAccumulator _acc;
            // Content-keyed dedup cache for literal (non-shared) Label cells GetString() can't serve
            // via the shared-string table; see ExcelReaderOptions.InternStrings.
            private readonly Utf8StringCache? _contentCache;
            private bool _ended;
            private int _row;

            internal Enumerator(XlsReader reader, int sheetOffset, CancellationToken ct = default)
            {
                _reader = reader;
                _ct = ct;
                _cursor = reader.OpenCursor(sheetOffset);
                _row = -1;
                _acc = new CellAccumulator(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes));
                _contentCache = reader._options.InternStrings ? new Utf8StringCache() : null;
            }

            /// <inheritdoc/>
            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan, rowBuffer: default, _reader.SharedStringCache, _contentCache);

            /// <inheritdoc/>
            public bool MoveNext()
            {
                _ct.ThrowIfCancellationRequested();
                return MoveNextCore();
            }

            /// <inheritdoc/>
            public ValueTask<bool> MoveNextAsync()
            {
                return new ValueTask<bool>(MoveNext());
            }

            private bool MoveNextCore()
            {
                // A row whose only records are BLANK/MULBLANK (styled empty cells — Excel writes these
                // routinely) yields zero cells; that must not end enumeration, so keep advancing to the
                // next row instead of returning false for anything short of true EOF.
                while (!_ended)
                {
                    ResetRow();
                    BiffCursor cursor = _cursor;
                    while (true)
                    {
                        long recordStart = cursor.Position;
                        if (!cursor.TryReadRecord(out int id, out ReadOnlySpan<byte> data))
                        {
                            _ended = true;
                            break;
                        }

                        if (!ReadRecord(cursor, recordStart, id, data))
                        {
                            break;
                        }
                    }
                    if (FinishRow())
                    {
                        return true;
                    }
                }
                return false;
            }

            // Returns false when the current row has ended or the worksheet stream reached EOF.
            //
            // A single switch below handles every cell-record kind: it reads the row (always at offset
            // 0), applies the row-tracking decision (first row / same row / new row started), and then
            // runs the type-specific parse in the same case, all in one dispatch on the record id.
            // Previously this was two separate switches on the same id — one just to test whether the
            // id carried a row, a second to parse it — twice the dispatch cost per cell record.
            private bool ReadRecord(BiffCursor cursor, long recordStart, int id, ReadOnlySpan<byte> data)
            {
                if (id == Rec.Bof)
                {
                    if (data.Length < 4 || ReadU16(data, 0) != Biff8Version || ReadU16(data, 2) != SubstreamWorksheet)
                    {
                        throw new NotSupportedException("Only BIFF8 worksheet streams are supported.");
                    }
                    return true;
                }
                if (id == Rec.Eof)
                {
                    _ended = true;
                    return false;
                }
                if (data.Length < 2)
                {
                    return true; // no room for the row field -- not a usable cell record, keep going
                }
                switch (id)
                {
                    case Rec.Label:
                        if (!AdvanceRow(cursor, recordStart, ReadU16(data, 0))) { return false; }
                        ParseLabel(data);
                        return true;
                    case Rec.LabelSst:
                        if (!AdvanceRow(cursor, recordStart, ReadU16(data, 0))) { return false; }
                        if (data.Length >= 10)
                        {
                            int col = ReadU16(data, 2);
                            int style = ReadU16(data, 4);
                            var (start, len, sharedIndex) = _reader.SharedAt(ReadI32(data, 6));
                            _acc.Add(col, start, len, CellType.ExcelString, style, CellValueSource.Shared, sharedIndex: sharedIndex);
                        }
                        return true;
                    case Rec.Number:
                        if (!AdvanceRow(cursor, recordStart, ReadU16(data, 0))) { return false; }
                        if (data.Length >= 14)
                        {
                            AddDouble(ReadU16(data, 2), ReadU16(data, 4), BinaryPrimitives.ReadDoubleLittleEndian(data.Slice(6, 8)));
                        }
                        return true;
                    case Rec.Rk:
                        if (!AdvanceRow(cursor, recordStart, ReadU16(data, 0))) { return false; }
                        if (data.Length >= 10)
                        {
                            AddDouble(ReadU16(data, 2), ReadU16(data, 4), Rk(ReadU32(data, 6)));
                        }
                        return true;
                    case Rec.MulRk:
                        if (!AdvanceRow(cursor, recordStart, ReadU16(data, 0))) { return false; }
                        ParseMulRk(data);
                        return true;
                    case Rec.BoolErr:
                        if (!AdvanceRow(cursor, recordStart, ReadU16(data, 0))) { return false; }
                        ParseBoolErr(data);
                        return true;
                    case Rec.Formula:
                        if (!AdvanceRow(cursor, recordStart, ReadU16(data, 0))) { return false; }
                        ParseFormula(data);
                        return true;
                    case Rec.Blank:
                    case Rec.MulBlank:
                        // Contributes no value, but still participates in row tracking.
                        return AdvanceRow(cursor, recordStart, ReadU16(data, 0));
                    default:
                        return true; // markup / unrecognized record -- keep going
                }
            }

            // Owns the row-tracking decision shared by every cell-record case above: the first cell
            // record of a row adopts its row number; a later record for the same row proceeds; a record
            // for a *different* row means the current row has ended, so the cursor rewinds to re-read
            // this same record as the first record of the next row.
            private bool AdvanceRow(BiffCursor cursor, long recordStart, int row)
            {
                if (_row < 0)
                {
                    _row = row;
                    return true;
                }
                if (row != _row)
                {
                    cursor.Position = recordStart;
                    return false;
                }
                return true;
            }

            private bool FinishRow()
            {
                _acc.SortByColumn();
                return _acc.Count > 0;
            }

            private void ResetRow()
            {
                _acc.Reset();
                _row = -1;
            }

            private void ParseLabel(ReadOnlySpan<byte> data)
            {
                if (data.Length < 9)
                {
                    return;
                }
                int col = ReadU16(data, 2);
                int style = ReadU16(data, 4);
                int chars = ReadU16(data, 6);
                byte flags = data[8];
                int valueStart = _acc.ValueLength;
                DecodeUnicodeString(data[9..], chars, flags);
                _acc.Add(col, valueStart, _acc.ValueLength - valueStart, CellType.ExcelString, style, CellValueSource.RowValues);
            }

            // Decodes an XLUnicodeString body (chars, already-read flags byte) that may continue
            // across one or more CONTINUE records — shared by LABEL and the FORMULA cached-string
            // result, both of which use this exact shape. Appends decoded UTF-8 to the accumulator.
            private void DecodeUnicodeString(ReadOnlySpan<byte> firstData, int chars, byte flags)
            {
                int firstByteLen = firstData.Length;
                int firstChars = (flags & 1) == 0 ? firstByteLen : firstByteLen / 2;
                firstChars = Math.Min(firstChars, chars);
                int firstBytes = (flags & 1) == 0 ? firstChars : firstChars * 2;
                Span<byte> dst = _acc.ReserveValueSpan((chars * 3) + 1);
                int written = DecodeStringToUtf8(firstData[..firstBytes], firstChars, flags, dst);
                _acc.Advance(written);
                int charsDecoded = firstChars;
                while (charsDecoded < chars && _cursor.PeekId() == Rec.Continue && _cursor.TryReadRecord(out _, out ReadOnlySpan<byte> cont) && cont.Length > 0)
                {
                    byte contFlags = cont[0];
                    int contByteLen = cont.Length - 1;
                    int contChars = (contFlags & 1) == 0 ? contByteLen : contByteLen / 2;
                    var remainingChars = chars - charsDecoded;
                    contChars = Math.Min(contChars, remainingChars);
                    int contBytes = (contFlags & 1) == 0 ? contChars : contChars * 2;
                    dst = _acc.ReserveValueSpan((remainingChars * 3) + 1);
                    written = DecodeStringToUtf8(cont.Slice(1, contBytes), contChars, contFlags, dst);
                    _acc.Advance(written);
                    charsDecoded += contChars;
                }
            }

            private void ParseMulRk(ReadOnlySpan<byte> data)
            {
                if (data.Length < 10)
                {
                    return;
                }
                int col = ReadU16(data, 2);
                int end = data.Length - 2;
                for (int pos = 4; pos + 6 <= end; pos += 6, col++)
                {
                    AddDouble(col, ReadU16(data, pos), Rk(ReadU32(data, pos + 2)));
                }
            }

            private void ParseBoolErr(ReadOnlySpan<byte> data)
            {
                if (data.Length < 8)
                {
                    return;
                }
                int col = ReadU16(data, 2);
                int style = ReadU16(data, 4);
                byte value = data[6];
                bool isError = data[7] != 0;
                if (isError)
                {
                    _acc.AddError(col, style, value);
                    return;
                }
                _acc.AddBool(col, style, value);
            }

            private void ParseFormula(ReadOnlySpan<byte> data)
            {
                if (data.Length < 20)
                {
                    return;
                }
                int col = ReadU16(data, 2);
                int style = ReadU16(data, 4);
                ReadOnlySpan<byte> result = data.Slice(6, 8);
                if (result[6] == 0xFF && result[7] == 0xFF)
                {
                    switch (result[0])
                    {
                        case 1:
                            _acc.AddBool(col, style, result[2]);
                            break;
                        case 2:
                            _acc.AddError(col, style, result[2]);
                            break;
                        case 0:
                            // String result: the marker means "see the STRING record that follows".
                            if (_cursor.PeekId() == Rec.StringRec && _cursor.TryReadRecord(out _, out ReadOnlySpan<byte> str) && str.Length >= 3)
                            {
                                int start = _acc.ValueLength;
                                int cch = ReadU16(str, 0);
                                byte strFlags = str[2];
                                DecodeUnicodeString(str[3..], cch, strFlags);
                                _acc.Add(col, start, _acc.ValueLength - start, CellType.ExcelString, style, CellValueSource.RowValues);
                            }
                            break;
                        default:
                            break;
                    }
                    return;
                }
                AddDouble(col, style, BinaryPrimitives.ReadDoubleLittleEndian(result), CellType.Formula);
            }

            private void AddDouble(int col, int style, double value, CellType? forced = null)
            {
                // Store the raw double only — no eager formatting. Cell formats lazily if a caller
                // asks for text (GetString/Value); numeric consumers read the double directly.
                CellType type = forced ?? (_reader.IsDateStyle(style) ? CellType.Date : CellType.Number);
                _acc.Add(col, _acc.ValueLength, 0, type, style, CellValueSource.RowValues, number: value, hasNumber: true);
            }

            /// <inheritdoc/>
            public void Dispose()
            {
                ReturnBuffers();
            }

            /// <inheritdoc/>
            public ValueTask DisposeAsync()
            {
                ReturnBuffers();
                return ValueTask.CompletedTask;
            }

            private void ReturnBuffers()
            {
                _cursor.Dispose();
                _acc.Return();
            }
        }
    }
}
