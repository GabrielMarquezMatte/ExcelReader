using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Writer
{
    // Adapts the flat CsvWriter to the ISheetWriter<TRow>/IWorkbookWriter<TSheet> pair that
    // WorkbookRecordWriter drives. A CSV file is a single sheet, so the workbook exposes exactly one
    // sheet and rejects a second AddSheet. The workbook owns the CsvWriter; the sheet only borrows it.
    public sealed class CsvSheetWriter : ISheetWriter<CsvRowWriter>
    {
        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "CsvWriter is borrowed; CsvWorkbookWriter owns and disposes it.")]
        private readonly CsvWriter _writer;

        internal CsvSheetWriter(CsvWriter writer)
        {
            _writer = writer;
        }

        public ValueTask StartAsync(CancellationToken ct = default)
        {
            return ValueTask.CompletedTask;
        }

        public ValueTask<CsvRowWriter> StartRowAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<CsvRowWriter>(_writer.StartRow());
        }

        public ValueTask EndAsync(CancellationToken ct = default)
        {
            return _writer.FlushAsync(ct);
        }

        public ValueTask DisposeAsync()
        {
            // The workbook owns the CsvWriter's lifetime; nothing to release here.
            return ValueTask.CompletedTask;
        }
    }

    public sealed class CsvWorkbookWriter : IWorkbookWriter<CsvSheetWriter>
    {
        private readonly CsvWriter _writer;
        private bool _sheetAdded;

        private CsvWorkbookWriter(CsvWriter writer)
        {
            _writer = writer;
        }

        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership of the CsvWriter transfers to CsvWorkbookWriter, which disposes it.")]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP004:Don't ignore created IDisposable",
            Justification = "The created CsvWriter is stored in the returned workbook's field and disposed there.")]
        public static CsvWorkbookWriter Create(Stream stream, bool leaveOpen = false, CsvWriterOptions? options = null)
        {
            return new CsvWorkbookWriter(CsvWriter.Create(stream, leaveOpen, options));
        }

        public ValueTask StartAsync(CancellationToken ct = default)
        {
            return ValueTask.CompletedTask;
        }

        public CsvSheetWriter AddSheet(string name)
        {
            if (_sheetAdded)
            {
                throw new InvalidOperationException("A CSV file holds a single sheet; only one AddSheet call is supported.");
            }
            _sheetAdded = true;
            return new CsvSheetWriter(_writer);
        }

        public ValueTask EndAsync(CancellationToken ct = default)
        {
            return _writer.FlushAsync(ct);
        }

        public ValueTask FlushAsync(CancellationToken ct = default)
        {
            return _writer.FlushAsync(ct);
        }

        public ValueTask DisposeAsync()
        {
            return _writer.DisposeAsync();
        }
    }
}
