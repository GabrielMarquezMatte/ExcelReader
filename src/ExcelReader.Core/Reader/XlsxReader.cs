using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsxReader : IDisposable, IAsyncDisposable
    {
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly ZipArchive _zip;
        private readonly (string Name, string Path)[] _sheets;
        private readonly bool[] _styleIsDate; // cellXfs index -> true when that style renders as a date/time
        private int _current;

        private byte[] _sharedFlat = [];      // pooled; all decoded shared-string bytes concatenated
        private int[] _sharedOffsets = [0];   // string i = _sharedFlat[_offsets[i].._offsets[i+1]]
        private int _sharedCount;
        private bool _sharedLoaded;

        // Sync open: reads the central directory and workbook/styles parts synchronously.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "Readonly field, first and only assignment in this constructor.")]
        internal XlsxReader(Stream stream, bool leaveOpen)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            try
            {
                _sheets = ParseSheets(Bytes(_zip, "xl/workbook.xml"), Bytes(_zip, "xl/_rels/workbook.xml.rels"));
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                _styleIsDate = ParseStyleDateFlags(Bytes(_zip, "xl/styles.xml"));
            }
            catch
            {
                _zip.Dispose();
                if (!leaveOpen)
                {
                    stream.Dispose();
                }
                throw;
            }
        }

        private XlsxReader(Stream stream, bool leaveOpen, ZipArchive zip,
            (string Name, string Path)[] sheets, bool[] styleIsDate)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _zip = zip;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
        }

        // Async open: central directory and parts are read with the .NET 10 async zip APIs.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "zip ownership transfers to the returned reader; disposed there or in the catch.")]
        internal static async ValueTask<XlsxReader> CreateAsync(Stream stream, bool leaveOpen, CancellationToken ct)
        {
            ZipArchive? zip = null;
            try
            {
                zip = await ZipArchive.CreateAsync(stream, ZipArchiveMode.Read, leaveOpen: true, entryNameEncoding: null, ct).ConfigureAwait(false);
                var wb = await BytesAsync(zip, "xl/workbook.xml", ct).ConfigureAwait(false);
                var rels = await BytesAsync(zip, "xl/_rels/workbook.xml.rels", ct).ConfigureAwait(false);
                var sheets = ParseSheets(wb, rels);
                if (sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                var styleIsDate = ParseStyleDateFlags(await BytesAsync(zip, "xl/styles.xml", ct).ConfigureAwait(false));
                return new XlsxReader(stream, leaveOpen, zip, sheets, styleIsDate);
            }
            catch
            {
                if (zip is not null)
                {
                    await zip.DisposeAsync().ConfigureAwait(false);
                }
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

        public string SheetName => _sheets[_current].Name;
        public int SheetCount => _sheets.Length;

        // A numeric cell whose style index maps to a date/time format is reported as CellType.Date.
        internal bool IsDateStyle(int style)
        {
            return (uint)style < (uint)_styleIsDate.Length && _styleIsDate[style];
        }


        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            for (int i = 0; i < _sheets.Length; i++)
            {
                if (name.Equals(_sheets[i].Name, StringComparison.OrdinalIgnoreCase))
                {
                    _current = i;
                    return true;
                }
            }
            return false;
        }

        public void MoveToSheet(int index)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _sheets.Length);
            _current = index;
        }

        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public Enumerator GetEnumerator()
        {
            EnsureSharedLoaded();
            var entry = _zip.GetEntry(_sheets[_current].Path)
                ?? throw new InvalidDataException($"Worksheet part not found: {_sheets[_current].Path}");
            return new Enumerator(this, entry.Open());
        }

        /// <summary>
        /// Streaming async enumerator over the current sheet. Use with a manual loop — <c>Current</c>
        /// is a ref struct (<c>Row</c>), so <c>await foreach</c> cannot bind it:
        /// <code>
        /// await using var e = await reader.GetAsyncEnumeratorAsync(ct);
        /// while (await e.MoveNextAsync()) { var row = e.Current; /* ... */ }
        /// </code>
        /// </summary>
        public async ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            await EnsureSharedLoadedAsync(ct).ConfigureAwait(false);
            var entry = _zip.GetEntry(_sheets[_current].Path)
                ?? throw new InvalidDataException($"Worksheet part not found: {_sheets[_current].Path}");
            var sheet = await entry.OpenAsync(ct).ConfigureAwait(false);
            return new Enumerator(this, sheet, ct);
        }

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal (int Start, int Length) SharedAt(int index)
        {
            if ((uint)index >= (uint)_sharedCount)
            {
                return (0, 0);
            }
            return (_sharedOffsets[index], _sharedOffsets[index + 1] - _sharedOffsets[index]);
        }

        public void Dispose()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            _zip.Dispose();
            if (!_leaveOpen)
            {
                _stream.Dispose();
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            await _zip.DisposeAsync().ConfigureAwait(false);
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }
    }
}
