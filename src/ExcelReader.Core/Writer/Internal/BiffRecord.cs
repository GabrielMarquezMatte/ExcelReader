namespace ExcelReader.Core.Writer.Internal
{
    // BIFF8 record type IDs and stream markers emitted by the writer (see [MS-XLS]).
    // The reader keeps its own private copy of the subset it consumes; sharing is deferred
    // to avoid churning the read path.
    internal static class BiffRecord
    {
        internal const int Bof = 0x0809;
        internal const int Eof = 0x000A;
        internal const int CodePage = 0x0042;
        internal const int Date1904 = 0x0022;
        internal const int Font = 0x0031;
        internal const int Format = 0x041E;
        internal const int Xf = 0x00E0;
        internal const int BoundSheet = 0x0085;
        internal const int Dimension = 0x0200;
        internal const int Number = 0x0203;
        internal const int Label = 0x0204;
        internal const int BoolErr = 0x0205;
        internal const int Blank = 0x0201;
        internal const int Continue = 0x003C;

        internal const int Biff8Version = 0x0600;
        internal const int SubstreamGlobals = 0x0005;
        internal const int SubstreamWorksheet = 0x0010;

        // Largest payload a single BIFF8 record can hold.
        internal const int MaxPayload = 8224;
    }
}
