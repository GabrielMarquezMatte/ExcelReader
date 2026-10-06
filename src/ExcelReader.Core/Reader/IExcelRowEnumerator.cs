namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// A forward-only cursor over a sheet's rows, implemented by every concrete format's row
    /// enumerator and driven either synchronously (<see cref="MoveNext"/>) or asynchronously
    /// (<see cref="MoveNextAsync"/>).
    /// </summary>
    public interface IExcelRowEnumerator : IDisposable, IAsyncDisposable
    {
        /// <summary>Gets the row at the enumerator's current position. Only valid after a call to <see cref="MoveNext"/> or <see cref="MoveNextAsync"/> has returned <see langword="true"/>.</summary>
        Row Current { get; }

        /// <summary>Advances the enumerator to the next row, reading synchronously.</summary>
        /// <returns><see langword="true"/> if a row was read; <see langword="false"/> if the sheet is exhausted.</returns>
        bool MoveNext();

        /// <summary>Advances the enumerator to the next row, reading asynchronously.</summary>
        /// <returns><see langword="true"/> if a row was read; <see langword="false"/> if the sheet is exhausted.</returns>
        ValueTask<bool> MoveNextAsync();
    }
}
