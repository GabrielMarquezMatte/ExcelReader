using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    // BIFF12 (.xlsb) reader. Uses the same ZIP/OPC container as .xlsx but worksheet parts are
    // binary BIFF12 records. The workbook, styles, and shared-string parts are read once at open
    // time (they're small); worksheets are streamed on demand by the enumerator.
    public sealed partial class XlsbReader : IExcelReader
    {
        // Shared-string pool: string i = _sharedFlat[_sharedOffsets[i].._sharedOffsets[i+1]].
        private readonly byte[] _sharedFlat = [];
        private readonly int[] _sharedOffsets = [0];
        private readonly bool[] _styleIsDate = [];

        [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Disposed in Dispose().")]
        private readonly ZipArchive? _zip;
        private readonly Stream? _stream;
        private readonly bool _leaveOpen;
        private readonly (string Name, string Path)[]? _sheets;
        private int _current;

        // Test-only: accepts pre-parsed components (no ZIP, no stream navigation).
        internal XlsbReader(byte[] sharedFlat, int[] sharedOffsets, bool[] styleIsDate, bool date1904)
        {
            _sharedFlat = sharedFlat;
            _sharedOffsets = sharedOffsets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
        }

        // Sync open: reads the three small workbook parts, keeps _zip open for worksheet streaming.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "Readonly field, first and only assignment in this constructor.")]
        internal XlsbReader(Stream stream, bool leaveOpen)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            try
            {
                var wb = ReadZipEntry(_zip, "xl/workbook.bin");
                _sheets = XlsbWorkbook.ParseSheets(wb, ReadZipEntry(_zip, "xl/_rels/workbook.bin.rels"));
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                _styleIsDate = XlsbStyles.ParseStyleDateFlags(ReadZipEntry(_zip, "xl/styles.bin"));
                IsDate1904 = XlsbWorkbook.ParseDate1904(wb);
                (_sharedFlat, _sharedOffsets) = XlsbSharedStrings.Parse(ReadZipEntry(_zip, "xl/sharedStrings.bin"));
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

        // Private constructor used by CreateAsync — all parts already parsed.
        private XlsbReader(Stream stream, bool leaveOpen, ZipArchive zip,
            (string Name, string Path)[] sheets, bool[] styleIsDate, bool date1904,
            byte[] sharedFlat, int[] sharedOffsets)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _zip = zip;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
            IsDate1904 = date1904;
            _sharedFlat = sharedFlat;
            _sharedOffsets = sharedOffsets;
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "zip ownership transfers to the returned reader on success, disposed in the catch on failure.")]
        internal static async ValueTask<XlsbReader> CreateAsync(Stream stream, bool leaveOpen, CancellationToken ct)
        {
            ZipArchive? zip = null;
            try
            {
#if NET10_0_OR_GREATER
                zip = await ZipArchive.CreateAsync(stream, ZipArchiveMode.Read, leaveOpen: true, entryNameEncoding: null, ct).ConfigureAwait(false);
#else
                ct.ThrowIfCancellationRequested();
                zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
#endif
                var wb = await ReadZipEntryAsync(zip, "xl/workbook.bin", ct).ConfigureAwait(false);
                var zipEntryData = await ReadZipEntryAsync(zip, "xl/_rels/workbook.bin.rels", ct).ConfigureAwait(false);
                var sheets = XlsbWorkbook.ParseSheets(wb, zipEntryData);
                if (sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                var styleIsDate = XlsbStyles.ParseStyleDateFlags(await ReadZipEntryAsync(zip, "xl/styles.bin", ct).ConfigureAwait(false));
                bool date1904 = XlsbWorkbook.ParseDate1904(wb);
                var (flat, offsets) = XlsbSharedStrings.Parse(await ReadZipEntryAsync(zip, "xl/sharedStrings.bin", ct).ConfigureAwait(false));
                return new XlsbReader(stream, leaveOpen, zip, sheets, styleIsDate, date1904, flat, offsets);
            }
            catch
            {
                if (zip is not null)
                {
#if NET10_0_OR_GREATER
                    await zip.DisposeAsync().ConfigureAwait(false);
#else
                    zip.Dispose();
#endif
                }
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

        // --- IExcelReader ---

        public bool IsDate1904 { get; }
        public string SheetName => _sheets![_current].Name;
        public int SheetCount => _sheets!.Length;

        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            for (int i = 0; i < _sheets!.Length; i++)
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
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _sheets!.Length);
            _current = index;
        }

        // --- Enumeration ---

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal bool IsDateStyle(int style) =>
            (uint)style < (uint)_styleIsDate.Length && _styleIsDate[style];

        internal (int Start, int Length) SharedAt(int index)
        {
            if ((uint)index >= (uint)(_sharedOffsets.Length - 1))
            {
                return (0, 0);
            }
            return (_sharedOffsets[index], _sharedOffsets[index + 1] - _sharedOffsets[index]);
        }

        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public Enumerator GetEnumerator()
        {
            var entry = _zip!.GetEntry(_sheets![_current].Path)
                ?? throw new InvalidDataException($"Worksheet part not found: {_sheets[_current].Path}");
            return new Enumerator(this, entry.Open());
        }

        public async ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            var entry = _zip!.GetEntry(_sheets![_current].Path)
                ?? throw new InvalidDataException($"Worksheet part not found: {_sheets[_current].Path}");
#if NET10_0_OR_GREATER
            Stream sheet = await entry.OpenAsync(ct).ConfigureAwait(false);
#else
            ct.ThrowIfCancellationRequested();
            Stream sheet = entry.Open();
#endif
            return new Enumerator(this, sheet, ct);
        }

        // Internal entry points used by Phase 3 tests — accept a pre-opened stream directly.
        internal Enumerator GetEnumerator(Stream sheetStream) => new(this, sheetStream);

        internal Enumerator GetAsyncEnumerator(Stream sheetStream, CancellationToken ct = default) =>
            new(this, sheetStream, ct);

        // --- Dispose ---

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_stream is disposed conditionally based on _leaveOpen.")]
        public void Dispose()
        {
            _zip?.Dispose(); // ZipArchive was opened with leaveOpen:true — does not close _stream
            if (!_leaveOpen)
            {
                _stream?.Dispose();
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "_stream is disposed conditionally based on _leaveOpen.")]
        public async ValueTask DisposeAsync()
        {
            if (_zip is not null)
            {
#if NET10_0_OR_GREATER
                await _zip.DisposeAsync().ConfigureAwait(false);
#else
                _zip.Dispose();
#endif
            }
            if (!_leaveOpen && _stream is not null)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }

        // --- ZIP helpers ---

        private static byte[] ReadZipEntry(ZipArchive zip, string name)
        {
            var entry = zip.GetEntry(name);
            if (entry is null)
            {
                return [];
            }
            var buf = new byte[entry.Length];
            using var s = entry.Open();
            s.ReadExactly(buf);
            return buf;
        }

        private static async ValueTask<byte[]> ReadZipEntryAsync(ZipArchive zip, string name, CancellationToken ct)
        {
            var entry = zip.GetEntry(name);
            if (entry is null)
            {
                return [];
            }
            var buf = new byte[entry.Length];
#if NET10_0_OR_GREATER
            Stream s = await entry.OpenAsync(ct).ConfigureAwait(false);
            await using (s.ConfigureAwait(false))
            {
                await s.ReadExactlyAsync(buf.AsMemory(), ct).ConfigureAwait(false);
            }
#else
            ct.ThrowIfCancellationRequested();
            using var s = entry.Open();
            await s.ReadExactlyAsync(buf.AsMemory(), ct).ConfigureAwait(false);
#endif
            return buf;
        }
    }
}
