using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        private const int MaxEncryptPasswordBytes = 4096;

        internal static int EncryptPackage(ReadOnlySpan<byte> packagePathUtf8, ReadOnlySpan<byte> destinationPathUtf8, ReadOnlySpan<byte> passwordUtf8)
        {
            ClearLastError();
            if (!IsValidEncryptPassword(passwordUtf8, "xl_encrypt_package"))
            {
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

        internal static int EncryptPackageToMemory(ReadOnlySpan<byte> package, ReadOnlySpan<byte> passwordUtf8, out byte[]? encrypted)
        {
            encrypted = null;
            ClearLastError();
            if (!IsValidEncryptPassword(passwordUtf8, "xl_encrypt_package_to_memory"))
            {
                return NativeStatus.InvalidArgument;
            }

            try
            {
                using MemoryStream destination = new();
                fixed (byte* data = package)
                {
                    using UnmanagedMemoryStream source = new(data, package.Length);
                    Excel.EncryptPackage(source, destination, Encoding.UTF8.GetString(passwordUtf8));
                }
                encrypted = destination.ToArray();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        private static bool IsValidEncryptPassword(ReadOnlySpan<byte> passwordUtf8, string export)
        {
            if (passwordUtf8.Length <= MaxEncryptPasswordBytes)
            {
                return true;
            }
            SetLastError($"{export} password_len must be at most {MaxEncryptPasswordBytes}; got {passwordUtf8.Length}.");
            return false;
        }
    }
}
