namespace ExcelReader.Core.Writer.Internal
{
    /// <summary>
    /// A file written beside its destination and moved over it only by <see cref="Commit"/>, so a failed,
    /// canceled or abandoned write leaves an existing destination untouched and removes the partial file.
    /// </summary>
    internal sealed class StagedFile : IDisposable, IAsyncDisposable
    {
        private readonly string _destination;
        private readonly string _temporary;
        private bool _committed;

        internal FileStream Stream { get; }

        internal StagedFile(string destination, bool useAsync = false)
        {
            _destination = Path.GetFullPath(destination);
            _temporary = Path.Combine(Path.GetDirectoryName(_destination) ?? string.Empty,
                $".excelreader-{Guid.NewGuid():N}.tmp");
            Stream = new FileStream(_temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                bufferSize: 4096, useAsync);
        }

        internal void Commit()
        {
            Stream.Dispose();
            File.Move(_temporary, _destination, overwrite: true);
            _committed = true;
        }

        internal async ValueTask CommitAsync()
        {
            await Stream.DisposeAsync().ConfigureAwait(false);
            File.Move(_temporary, _destination, overwrite: true);
            _committed = true;
        }

        public void Dispose()
        {
            if (_committed)
            {
                return;
            }
            FailureCleanup.Dispose(Stream);
            File.Delete(_temporary);
        }

        public async ValueTask DisposeAsync()
        {
            if (_committed)
            {
                return;
            }
            await FailureCleanup.DisposeAsync(Stream).ConfigureAwait(false);
            File.Delete(_temporary);
        }
    }
}
