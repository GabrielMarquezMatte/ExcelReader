using System.Buffers.Binary;
using System.Text;

namespace ExcelReader.Tests
{
    // Hand-builds BIFF12 (.xlsb) record bytes for parser tests: record framing (varint id + len)
    // plus the field encoders that mirror the decoders in ExcelReader.Core.Reader.Biff12.
    internal static class Biff12Build
    {
        internal static byte[] Record(int id, params byte[] payload)
        {
            return [.. Id(id), .. Len(payload.Length), .. payload];
        }

        internal static byte[] U16(int value)
        {
            byte[] bytes = new byte[2];
            BinaryPrimitives.WriteUInt16LittleEndian(bytes, (ushort)value);
            return bytes;
        }

        internal static byte[] U32(uint value)
        {
            byte[] bytes = new byte[4];
            BinaryPrimitives.WriteUInt32LittleEndian(bytes, value);
            return bytes;
        }

        internal static byte[] F64(double value)
        {
            byte[] bytes = new byte[8];
            BinaryPrimitives.WriteDoubleLittleEndian(bytes, value);
            return bytes;
        }

        // XLWideString: cch (u32) + UTF-16LE chars.
        internal static byte[] WideString(string value)
        {
            return [.. U32((uint)value.Length), .. Encoding.Unicode.GetBytes(value)];
        }

        // XLNullableWideString null sentinel.
        internal static byte[] NullWideString()
        {
            return U32(0xFFFFFFFF);
        }

        // A 16-byte BrtXF payload with the given numFmtId at offset 2 (ixfeParent at 0).
        internal static byte[] Xf(int numFmtId)
        {
            return [.. U16(0), .. U16(numFmtId), .. new byte[12]];
        }

        // Cell record payloads: col(u32) + styleAndFlags(u32) + value.
        internal static byte[] CellRk(int col, int style, uint rk)
        {
            return [.. U32((uint)col), .. U32((uint)style), .. U32(rk)];
        }

        internal static byte[] CellReal(int col, int style, double value)
        {
            return [.. U32((uint)col), .. U32((uint)style), .. F64(value)];
        }

        internal static byte[] CellIsst(int col, int style, uint isst)
        {
            return [.. U32((uint)col), .. U32((uint)style), .. U32(isst)];
        }

        internal static byte[] CellSt(int col, int style, string value)
        {
            return [.. U32((uint)col), .. U32((uint)style), .. WideString(value)];
        }

        internal static byte[] CellBool(int col, int style, bool value)
        {
            return [.. U32((uint)col), .. U32((uint)style), value ? (byte)1 : (byte)0];
        }

        internal static byte[] CellError(int col, int style, byte error)
        {
            return [.. U32((uint)col), .. U32((uint)style), error];
        }

        // BrtCellRString: col(u32) + styleAndFlags(u32) + cRun(byte) + XLWideString.
        internal static byte[] CellRString(int col, int style, byte cRun, string value)
        {
            return [.. U32((uint)col), .. U32((uint)style), cRun, .. WideString(value)];
        }


        private static byte[] Id(int id)
        {
            if (id < 0x80)
            {
                return [(byte)id];
            }
            return [(byte)((id & 0x7F) | 0x80), (byte)(id >> 7)];
        }

        private static byte[] Len(int value)
        {
            List<byte> bytes = [];
            do
            {
                byte b = (byte)(value & 0x7F);
                value >>= 7;
                if (value != 0)
                {
                    b |= 0x80;
                }
                bytes.Add(b);
            }
            while (value != 0);
            return [.. bytes];
        }
    }
}
