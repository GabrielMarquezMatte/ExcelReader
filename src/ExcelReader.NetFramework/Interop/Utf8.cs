using System.Text;

namespace ExcelReader.NetFramework.Interop
{
    internal static class Utf8
    {
        internal static byte[] Bytes(string s)
        {
            return Encoding.UTF8.GetBytes(s);
        }

        internal static unsafe string? Decode(byte* p, int len)
        {
            if (len == 0)
            {
                return null;
            }

            return Encoding.UTF8.GetString(p, len);
        }
    }
}
