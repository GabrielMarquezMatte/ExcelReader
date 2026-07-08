namespace ExcelReader.Core.Writer.Internal
{
    // BIFF8 record type IDs and stream markers emitted by the writer (see [MS-XLS]).
    // The reader keeps its own private copy of the subset it consumes; sharing is deferred
    // to avoid churning the read path.
    internal static class BiffRecord
    {
        internal const int Bof = 0x0809;
        internal const int Eof = 0x000A;
        internal const int InterfaceHdr = 0x00E1;
        internal const int Mms = 0x00C1;
        internal const int InterfaceEnd = 0x00E2;
        internal const int WriteAccess = 0x005C;
        internal const int CodePage = 0x0042;
        internal const int Dsf = 0x0161;
        internal const int Date1904 = 0x0022;
        internal const int Font = 0x0031;
        internal const int Xf = 0x00E0;
        internal const int Style = 0x0293;
        internal const int BoundSheet = 0x0085;
        internal const int TabId = 0x013D;
        internal const int FnGroupCount = 0x009C;
        internal const int Window1 = 0x003D;
        internal const int Backup = 0x0040;
        internal const int HideObj = 0x008D;
        internal const int Precision = 0x000E;
        internal const int RefreshAll = 0x01B7;
        internal const int BookBool = 0x00DA;
        internal const int Dimension = 0x0200;
        internal const int Window2 = 0x023E;
        internal const int Number = 0x0203;
        internal const int Label = 0x0204;
        internal const int BoolErr = 0x0205;

        internal const int Biff8Version = 0x0600;
        internal const int SubstreamGlobals = 0x0005;
        internal const int SubstreamWorksheet = 0x0010;

        // Largest payload a single BIFF8 record can hold.
        internal const int MaxPayload = 8224;
    }
}
