namespace ExcelReader.Core.Reader
{
    internal ref struct Biff12RecordReader
    {
        private readonly ReadOnlySpan<byte> _data;

        internal Biff12RecordReader(ReadOnlySpan<byte> data)
        {
            _data = data;
            Position = 0;
        }

        internal int Position { get; private set; }

        internal bool TryReadRecord(out int id, out ReadOnlySpan<byte> payload)
        {
            ReadOnlySpan<byte> data = _data;
            int fast = Position;
            if (fast + 1 >= data.Length)
            {
                return TryReadRecordSlow(out id, out payload);
            }
            int b0 = data[fast];
            int len = data[fast + 1];
            if (((b0 | len) & 0x80) != 0 || fast + 2 + len > data.Length)
            {
                return TryReadRecordSlow(out id, out payload);
            }
            id = b0;
            payload = data.Slice(fast + 2, len);
            Position = fast + 2 + len;
            return true;
        }

        private bool TryReadRecordSlow(out int id, out ReadOnlySpan<byte> payload)
        {
            id = 0;
            payload = default;
            int pos = Position;
            if (!TryReadId(ref pos, out id))
            {
                return false;
            }
            if (!TryReadLength(ref pos, out int length))
            {
                return false;
            }
            if (pos + length > _data.Length)
            {
                return false;
            }
            payload = _data.Slice(pos, length);
            Position = pos + length;
            return true;
        }

        private readonly bool TryReadId(ref int pos, out int id)
        {
            id = 0;
            if (pos >= _data.Length)
            {
                return false;
            }
            byte b0 = _data[pos];
            if ((b0 & 0x80) == 0)
            {
                id = b0;
                pos += 1;
                return true;
            }
            if (pos + 1 >= _data.Length)
            {
                return false;
            }
            id = (b0 & 0x7F) | (_data[pos + 1] << 7);
            pos += 2;
            return true;
        }

        private readonly bool TryReadLength(ref int pos, out int length)
        {
            length = 0;
            for (int shift = 0; shift < 28; shift += 7)
            {
                if (pos >= _data.Length)
                {
                    length = 0;
                    return false;
                }
                byte b = _data[pos++];
                length |= (b & 0x7F) << shift;
                if ((b & 0x80) == 0)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
