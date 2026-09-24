using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Xlsb
{
    public sealed partial class XlsbReader
    {
        /// <summary>Forward-only enumerator over an <see cref="XlsbReader"/> sheet's rows.</summary>
        /// <remarks>
        /// Streams the underlying binary <c>sheetN.bin</c> entry through a refillable pooled buffer;
        /// <c>Biff12RecordReader</c> framing guarantees that a partial record at the buffer boundary
        /// is detected and retried after the next fill.
        /// </remarks>
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : PooledStreamRowEnumerator, IExcelRowEnumerator
        {
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
                Justification = "XlsbReader is borrowed; its lifetime is managed by the caller, not this enumerator.")]
            private readonly XlsbReader _reader;
            private readonly bool[] _styleIsDate;
            private readonly int[] _sharedOffsets;
            private readonly Utf8StringCache? _contentCache;
            private readonly ZipArchiveEntry? _entry;
            private bool _ended;
            private bool _pendingRowHdr;

            internal Enumerator(XlsbReader reader, Stream sheet, long entryLength = 0, CancellationToken ct = default)
                : base(sheet, reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), WorkbookLookups.InitialBufferCapacity(entryLength), ownsSource: true, ct)
            {
                _reader = reader;
                _styleIsDate = reader._styleIsDate;
                _sharedOffsets = reader._sharedOffsets;
                _contentCache = reader._options.InternStrings ? new Utf8StringCache() : null;
            }

            internal Enumerator(XlsbReader reader, ZipArchiveEntry entry, CancellationToken ct)
                : base(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), WorkbookLookups.InitialBufferCapacity(entry.Length), ct)
            {
                _reader = reader;
                _styleIsDate = reader._styleIsDate;
                _sharedOffsets = reader._sharedOffsets;
                _contentCache = reader._options.InternStrings ? new Utf8StringCache() : null;
                _entry = entry;
            }

            private protected override Stream OpenSource()
            {
                return WorkbookLookups.OpenEntryStream(_entry!, _reader._decompressedBytes, _reader._options);
            }

            private protected override async ValueTask<Stream> OpenSourceAsync()
            {
                return await WorkbookLookups.OpenEntryStreamAsync(_entry!, _reader._decompressedBytes, _reader._options, _ct).ConfigureAwait(false);
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
                _ct.ThrowIfCancellationRequested();
                if (_ended)
                {
                    return new ValueTask<bool>(false);
                }
                while (true)
                {
                    ResetRow();
                    if (!_pendingRowHdr)
                    {
                        int seek = SeekRowHdrFromBuffer();
                        if (seek == 0)
                        {
                            return new ValueTask<bool>(false);
                        }
                        if (seek == 2)
                        {
                            return MoveNextRowAsync(seekDone: false);
                        }
                    }
                    _pendingRowHdr = false;
                    int collect = CollectCellsFromBuffer();
                    if (collect == 2)
                    {
                        return MoveNextRowAsync(seekDone: true);
                    }
                    if (_acc.Count > 0)
                    {
                        return new ValueTask<bool>(true);
                    }
                    if (!_pendingRowHdr)
                    {
                        return new ValueTask<bool>(false);
                    }
                }
            }

            private async ValueTask<bool> MoveNextRowAsync(bool seekDone)
            {
                if (!seekDone && !await SeekRowHdrAsync().ConfigureAwait(false))
                {
                    return false;
                }
                _pendingRowHdr = false;
                await CollectCellsAsync().ConfigureAwait(false);
                while (true)
                {
                    if (_acc.Count > 0)
                    {
                        return true;
                    }
                    if (!_pendingRowHdr)
                    {
                        return false;
                    }
                    ResetRow();
                    if (!await SeekRowHdrAsync().ConfigureAwait(false))
                    {
                        return false;
                    }
                    _pendingRowHdr = false;
                    await CollectCellsAsync().ConfigureAwait(false);
                }
            }

            private bool MoveNextCore()
            {
                if (_ended)
                {
                    return false;
                }
                while (true)
                {
                    ResetRow();
                    if (!_pendingRowHdr && !SkipToRowHdr())
                    {
                        return false;
                    }
                    _pendingRowHdr = false;
                    CollectCells();
                    if (_acc.Count > 0)
                    {
                        return true;
                    }
                    if (!_pendingRowHdr)
                    {
                        return false;
                    }
                }
            }

            private bool SkipToRowHdr()
            {
                while (true)
                {
                    int result = SeekRowHdrFromBuffer();
                    if (result != 2)
                    {
                        return result == 1;
                    }
                    Fill();
                }
            }

            private void CollectCells()
            {
                while (CollectCellsFromBuffer() == 2)
                {
                    Fill();
                }
            }

            private async ValueTask<bool> SeekRowHdrAsync()
            {
                while (true)
                {
                    int result = SeekRowHdrFromBuffer();
                    if (result == 1)
                    {
                        return true;
                    }
                    if (result == 0)
                    {
                        return false;
                    }
                    await FillAsync().ConfigureAwait(false);
                }
            }

            private async ValueTask CollectCellsAsync()
            {
                while (true)
                {
                    int result = CollectCellsFromBuffer();
                    if (result != 2)
                    {
                        return;
                    }
                    await FillAsync().ConfigureAwait(false);
                }
            }

            private int SeekRowHdrFromBuffer()
            {
                Biff12RecordReader reader = new(_buf.AsSpan(_pos, _len - _pos));
                while (reader.TryReadRecord(out int id, out _))
                {
                    if (id == Brt.RowHdr)
                    {
                        _pos += reader.Position;
                        return 1;
                    }
                    if (IsEndSheetData(id))
                    {
                        _pos += reader.Position;
                        _ended = true;
                        return 0;
                    }
                }
                _pos += reader.Position;
                if (_eof)
                {
                    ThrowIfTruncated();
                    _ended = true;
                    return 0;
                }
                return 2;
            }

            private int CollectCellsFromBuffer()
            {
                var reader = new Biff12RecordReader(_buf.AsSpan(_pos, _len - _pos));
                while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
                {
                    if (id == Brt.RowHdr)
                    {
                        _pos += reader.Position;
                        _pendingRowHdr = true;
                        return 1;
                    }
                    if (IsEndSheetData(id))
                    {
                        _pos += reader.Position;
                        _ended = true;
                        return 0;
                    }
                    ProcessCell(id, payload);
                }
                _pos += reader.Position;
                if (_eof)
                {
                    ThrowIfTruncated();
                    _ended = true;
                    return 0;
                }
                return 2;
            }

            private void ProcessCell(int id, ReadOnlySpan<byte> payload)
            {
                if (payload.Length < 8)
                {
                    return;
                }
                int col = (int)Biff12.ReadU32(payload, 0);
                int style = (int)(Biff12.ReadU32(payload, 4) & 0x00FFFFFF);

                switch (id)
                {
                    case Brt.CellRk when payload.Length >= 12:
                        AddDouble(col, style, Biff12.Rk(Biff12.ReadU32(payload, 8)));
                        break;
                    case Brt.CellReal or Brt.FmlaNum when payload.Length >= 16:
                        AddDouble(col, style, Biff12.ReadF64(payload, 8));
                        break;
                    case Brt.CellIsst when payload.Length >= 12:
                        var (start, len, sharedIndex) = WorkbookLookups.SharedAt(_sharedOffsets, (int)Biff12.ReadU32(payload, 8));
                        _acc.Add(col, start, len, CellType.ExcelString, style, CellValueSource.Shared, sharedIndex: sharedIndex);
                        break;

                    case Brt.CellSt or Brt.FmlaString:
                        AddInlineString(col, style, payload, 8);
                        break;
                    case Brt.CellBool or Brt.FmlaBool when payload.Length >= 9:
                        AppendBool(col, style, payload[8]);
                        break;
                    case Brt.CellError or Brt.FmlaError when payload.Length >= 9:
                        AppendError(col, style, payload[8]);
                        break;
                    case Brt.CellRString when payload.Length >= 9:
                        AddInlineString(col, style, payload, 9);
                        break;
                }
            }

            [MethodImpl(MethodImplOptions.NoInlining)]
            private void AddInlineString(int col, int style, ReadOnlySpan<byte> payload, int offset)
            {
                if (Biff12.TryReadWideString(payload, offset, out ReadOnlySpan<char> chars, out _))
                {
                    AppendString(col, style, chars);
                }
            }

            private static bool IsEndSheetData(int id)
            {
                return id == Brt.EndSheetData;
            }

            private void AddDouble(int col, int style, double value)
            {
                CellType type = WorkbookLookups.IsDateStyle(_styleIsDate, style) ? CellType.Date : CellType.Number;
                CellAccumulator acc = _acc;
                acc.Add(col, acc.ValueLength, 0, type, style, CellValueSource.RowValues, number: value, hasNumber: true);
            }

            private void AppendString(int col, int style, ReadOnlySpan<char> chars)
            {
                int start = _acc.ValueLength;
                Span<byte> dst = _acc.ReserveValueSpan(Encoding.UTF8.GetMaxByteCount(chars.Length));
                _acc.Advance(Encoding.UTF8.GetBytes(chars, dst));
                _acc.Add(col, start, _acc.ValueLength - start, CellType.ExcelString, style, CellValueSource.RowValues);
            }

            private void AppendBool(int col, int style, byte value)
            {
                _acc.AddBool(col, style, value);
            }

            private void AppendError(int col, int style, byte error)
            {
                _acc.AddError(col, style, error);
            }

            private void ResetRow()
            {
                _acc.Reset();
            }


            private void ThrowIfTruncated()
            {
                if (_pos < _len)
                {
                    throw new InvalidDataException(
                        $"Truncated XLSB worksheet stream: {_len - _pos} trailing byte(s) do not form a complete record.");
                }
            }

        }
    }
}
