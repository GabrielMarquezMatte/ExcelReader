namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// A forward-only cursor over a sheet's rows, implemented by every concrete format's row
    /// enumerator and driven either synchronously (<see cref="MoveNext"/>) or asynchronously
    /// (<see cref="MoveNextAsync"/>).
    /// </summary>
    /// <remarks>Only ExcelReader implements this interface, and a minor version can add members to it. See <see href="https://gabrielmarquezmatte.github.io/ExcelReader/docs/guide/versioning.html">versioning and compatibility</see>.</remarks>
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
