namespace ExcelReader.Native
{
    internal static class NativeStatus
    {
        internal const int Ok = 0;
        internal const int Eof = -1;
        internal const int BufferTooSmall = -2;
        internal const int InvalidHandle = -3;
        internal const int InvalidArgument = -4;
        internal const int Error = -5;

        internal const int PasswordRequired = -6;
        internal const int PasswordIncorrect = -7;

        internal const int AbiVersion = 5;
    }

    internal static class NativeLimits
    {
        internal const int MaxColumnSpecs = 16_384;

        internal const int MaxColumnNameBytes = 131_068;

        internal const int MaxNamesPerSpec = 32;
    }

    internal static class NativeFormat
    {
        internal const int Auto = 0;
        internal const int Xls = 1;
        internal const int Xlsx = 2;
        internal const int Xlsb = 3;
        internal const int Csv = 4;
    }
}
