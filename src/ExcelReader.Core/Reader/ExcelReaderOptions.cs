using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Resource limits and behavior toggles applied when opening and reading a workbook or a CSV source.
    /// Exceeding a byte or entry-count limit throws <see cref="ExcelLimitExceededException"/>.
    /// </summary>
    public sealed record ExcelReaderOptions
    {
        /// <summary>Gets the maximum total decompressed bytes allowed for one read of a sheet. Defaults to 512 MiB.</summary>
        /// <remarks>Applies to ZIP-backed formats (XLSX/XLSB) as their decompressed byte budget: what the
        /// workbook inflated when it opened and loaded its shared strings, plus the sheet being read. Each
        /// read of a sheet starts from that one-time total again, so reading a sheet twice does not count
        /// it twice. For the legacy CFB container (.xls) it caps the declared Workbook stream size — the
        /// CFB path has nothing to decompress, but this is still the caller's budget for what that phase
        /// may materialize.</remarks>
        public long MaxTotalDecompressedBytes { get; init; } = 512L * 1024 * 1024;

        /// <summary>Gets the maximum byte length allowed for a single cell's value, or a single buffered CSV record. Defaults to 32 MiB.</summary>
        public int MaxCellBytes { get; init; } = 32 * 1024 * 1024;

        /// <summary>Gets the maximum total byte size allowed for the shared string table. Defaults to 128 MiB.</summary>
        public long MaxSharedStringBytes { get; init; } = 128L * 1024 * 1024;

        /// <summary>Gets the maximum number of ZIP entries allowed in an XLSX/XLSB archive. Defaults to 65,536.</summary>
        public int MaxZipEntries { get; init; } = 65_536;

        /// <summary>Gets a value indicating whether ZIP-backed sheet data is decompressed on a
        /// background thread ahead of parsing. Defaults to <see langword="false"/>.</summary>
        /// <remarks>Applies only to ZIP-backed formats (XLSX/XLSB); XLS and CSV have nothing to
        /// decompress, so the option is silently ignored for them. Intended for single-file batch
        /// processing, where overlapping inflate with parsing shortens one read's wall-clock time —
        /// not for concurrent server workloads, where the extra background thread per read competes
        /// with work already saturating the CPU.</remarks>
        public bool PrefetchDecompression { get; init; }

        /// <summary>Gets a value indicating whether each ZIP entry read from an XLSX or XLSB package is
        /// checked against the CRC-32 its central directory records. Defaults to <see langword="true"/>.</summary>
        /// <remarks>A mismatch throws <see cref="InvalidDataException"/> when the entry's last bytes are read, so
        /// rows before the damage may already have been returned. Turn it off to read a package whose
        /// producer wrote wrong checksums; damaged data then surfaces only if it also breaks parsing.</remarks>
        public bool VerifyZipChecksums { get; init; } = true;

        /// <summary>
        /// Gets a value indicating whether text cells not served by a shared-string table (inline and
        /// formula-string cells, and every CSV field) are deduplicated through a small content-keyed cache
        /// when materialized via <see cref="Cell.GetString"/>. Defaults to <see langword="false"/>.
        /// Has no effect on shared-string cells, which already dedup through the shared-string table's
        /// own index-keyed cache, nor on the zero-copy <see cref="Cell.Value"/> span path —
        /// only <c>GetString()</c> consults this cache. Worth enabling only for genuinely categorical
        /// (low-cardinality, repeated) text; for high-cardinality data every lookup is a cache miss, so the
        /// per-call hashing cost is pure overhead — measured roughly doubling wall-clock time on a
        /// real-world, mixed-cardinality CSV corpus.
        /// </summary>
        public bool InternStrings { get; init; }

        /// <summary>Gets the password used to open an encrypted workbook. Defaults to <see langword="null"/>
        /// (no password), in which case opening an encrypted workbook throws
        /// <see cref="ExcelEncryptionException"/> with <see cref="ExcelEncryptionReason.PasswordRequired"/>.</summary>
        /// <remarks>Applies to encrypted OOXML workbooks (.xlsx/.xlsb/.xlsm). Encrypted legacy .xls files
        /// are not supported at all, with or without a password.</remarks>
        public ExcelPassword? Password { get; init; }

        /// <summary>Gets a value indicating whether an encrypted workbook's <c>dataIntegrity</c> HMAC is
        /// verified when reading from a stream. Defaults to <see langword="false"/>.</summary>
        /// <remarks>The HMAC covers the whole encrypted package, so verifying it on the streaming path
        /// requires a full pass over the file before the first row — making time-to-first-row proportional
        /// to file size, which defeats the point of a streaming reader. The in-memory path
        /// (<see cref="Excel.FromXlsx(ReadOnlyMemory{byte},ExcelReaderOptions?)"/>) has already decrypted
        /// everything, so it always verifies regardless of this setting. With verification off, targeted
        /// ciphertext tampering is not detected, though corrupt ciphertext still fails loudly downstream
        /// when the decrypted bytes fail ZIP and XML parsing. Standard encryption has no HMAC field, so
        /// this setting does not apply to it.</remarks>
        public bool VerifyEncryptedIntegrity { get; init; }

        /// <summary>Gets the maximum password-derivation iteration count accepted from an encrypted
        /// workbook. Defaults to 100,000, which is what Excel writes.</summary>
        /// <remarks>An agile-encrypted file states its own iteration count, so a crafted workbook can ask
        /// for billions of iterations and stall the process for hours. Exceeding this throws
        /// <see cref="ExcelLimitExceededException"/>. ECMA-376 standard encryption is unaffected: its
        /// 50,000 iterations are fixed by the scheme rather than read from the file, so there is no
        /// file-supplied value to bound and this limit is not applied on that path.</remarks>
        public int MaxPasswordSpinCount { get; init; } = 100_000;

        /// <summary>Gets the delimited-text settings used when a source is opened as
        /// <see cref="ExcelFileFormat.Csv"/>. Defaults to <see langword="null"/>, meaning
        /// <see cref="CsvReaderOptions.Default"/>.</summary>
        /// <remarks>Ignored for the Excel formats. On the CSV path, <see cref="MaxCellBytes"/> and
        /// <see cref="InternStrings"/> still apply; the ZIP, shared-string, encryption and prefetch
        /// settings have nothing to act on.</remarks>
        public CsvReaderOptions? Csv { get; init; }

        /// <summary>Gets the default options instance, used whenever a reader is opened without explicit options.</summary>
        public static ExcelReaderOptions Default { get; } = new();
    }
}
