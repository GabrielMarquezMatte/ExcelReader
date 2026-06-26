using System.Text;

namespace ExcelReader.Core.Reader
{
    // Decodes xl/sharedStrings.bin (binary BIFF12) into a flat UTF-8 buffer + offsets, matching how
    // XlsxReader stores shared strings so Cell consumes UTF-8 either way:
    // string i = Flat[Offsets[i]..Offsets[i+1]].
    internal static class XlsbSharedStrings
    {
        internal static (byte[] Flat, int[] Offsets) Parse(ReadOnlySpan<byte> sharedBin)
        {
            if (sharedBin.IsEmpty)
            {
                return ([], [0]);
            }
            // UTF-8 is at most 1.5x the UTF-16 byte size; the part size is a fine starting hint and
            // the buffer grows if a run of multi-byte characters overruns it.
            byte[] flat = new byte[Math.Max(16, sharedBin.Length)];
            int flatLen = 0;
            List<int> offsets = [0];
            var reader = new Biff12RecordReader(sharedBin);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                if (id != Brt.SSTItem)
                {
                    continue;
                }
                flatLen = AppendItem(payload, ref flat, flatLen);
                offsets.Add(flatLen);
            }
            return (flat, [.. offsets]);
        }

        // BrtSSTItem is a RichStr: 1 flags byte, then the text as a wide string; trailing rich/phonetic
        // runs come after the text and are ignored (the record framing already bounds them).
        private static int AppendItem(ReadOnlySpan<byte> payload, ref byte[] flat, int flatLen)
        {
            if (payload.Length < 1 || !Biff12.TryReadWideString(payload, 1, out ReadOnlySpan<char> chars, out _))
            {
                return flatLen;
            }
            int needed = flatLen + Encoding.UTF8.GetByteCount(chars);
            if (needed > flat.Length)
            {
                Array.Resize(ref flat, Math.Max(flat.Length * 2, needed));
            }
            return flatLen + Encoding.UTF8.GetBytes(chars, flat.AsSpan(flatLen));
        }
    }
}
