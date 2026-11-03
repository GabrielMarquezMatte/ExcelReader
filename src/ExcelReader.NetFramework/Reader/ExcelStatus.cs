using ExcelReader.NetFramework.Parser.Internal;

namespace ExcelReader.NetFramework.Reader
{
    /// <summary>Status codes returned by the native library.</summary>
    public enum ExcelStatus
    {
        /// <summary>The call succeeded.</summary>
        Ok = 0,
        /// <summary>No more rows.</summary>
        Eof = -1,
        /// <summary>The supplied buffer was too small.</summary>
        BufferTooSmall = -2,
        /// <summary>The handle is closed or unknown.</summary>
        InvalidHandle = -3,
        /// <summary>An argument was rejected.</summary>
        InvalidArgument = -4,
        /// <summary>The call failed.</summary>
        Error = -5,
        /// <summary>The workbook is encrypted and no password was supplied.</summary>
        PasswordRequired = -6,
        /// <summary>The supplied password did not match the workbook.</summary>
        PasswordIncorrect = -7,
    }
}
