namespace ExcelReader.Core.Reader
{
    // Decodes the BIFF12 record framing used by .xlsb parts: a variable-length record id
    // (1–2 bytes, 7 bits each, high bit = continuation) followed by a variable-length payload
    // size (1–4 bytes, same varint scheme), then the payload. Forward-only over an in-memory
    // span; the worksheet enumerator feeds it a buffer window holding whole records.
    internal ref struct Biff12RecordReader
    {
        private readonly ReadOnlySpan<byte> _data;

        internal Biff12RecordReader(ReadOnlySpan<byte> data)
        {
            _data = data;
            Position = 0;
        }

        internal int Position { get; private set; }

        // Reads the next (id, payload). Returns false at end of data, or when a record's framing
        // or payload would run past the end — leaving Position unchanged, so a streaming caller can
        // refill its buffer and retry from the same record.
        internal bool TryReadRecord(out int id, out ReadOnlySpan<byte> payload)
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
            // Two-byte id: low 7 bits of b0, then the next byte shifted up. Ids fit in 14 bits.
            id = (b0 & 0x7F) | (_data[pos + 1] << 7);
            pos += 2;
            return true;
        }

        private readonly bool TryReadLength(ref int pos, out int length)
        {
            length = 0;
            // Up to 4 bytes, 7 bits each (28 bits total — well within int).
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
            return true;
        }
    }
}
