using System.Text;

namespace ExcelReader.NetFramework.Interop
{
    internal static class Utf8
    {
        internal static byte[] Bytes(string s)
        {
            return Encoding.UTF8.GetBytes(s);
        }

        internal static byte[] NonNullBytes(string s, out int length)
        {
            length = Encoding.UTF8.GetByteCount(s);
            byte[] bytes = new byte[Math.Max(length, 1)];
            Encoding.UTF8.GetBytes(s, 0, s.Length, bytes, 0);
            return bytes;
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
