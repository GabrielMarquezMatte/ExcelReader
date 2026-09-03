using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        // Same ceiling as NativeOpenOptions's password field - the two are the same class of
        // caller-supplied length driving an allocation, so they share a bound.
        private const int MaxEncryptPasswordBytes = 4096;

        internal static int EncryptPackage(ReadOnlySpan<byte> packagePathUtf8, ReadOnlySpan<byte> destinationPathUtf8, ReadOnlySpan<byte> passwordUtf8)
        {
            ClearLastError();
            if (passwordUtf8.Length > MaxEncryptPasswordBytes)
            {
                SetLastError($"xl_encrypt_package password_len must be at most {MaxEncryptPasswordBytes}; got {passwordUtf8.Length}.");
                return NativeStatus.InvalidArgument;
            }

            try
            {
                string packagePath = Encoding.UTF8.GetString(packagePathUtf8);
                string destinationPath = Encoding.UTF8.GetString(destinationPathUtf8);
                string password = Encoding.UTF8.GetString(passwordUtf8);
                Excel.EncryptPackage(packagePath, destinationPath, password);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }
    }
}
