using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Xlsb
{
    internal static class XlsbTables
    {
        private const int BeginListFixedSize = 64;
        private const int BeginListColFixedSize = 24;
        private const int StyleClientFlagsSize = 2;

        internal static TablePart ParseTable(int sheetIndex, ReadOnlySpan<byte> part)
        {
            ListHeader? header = null;
            List<string> columns = [];
            string? styleName = null;
            var reader = new Biff12RecordReader(part);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                switch (id)
                {
                    case Brt.BeginList:
                        header = ReadBeginList(payload);
                        break;
                    case Brt.BeginListCol:
                        columns.Add(ReadColumnName(payload));
                        break;
                    case Brt.ListTableStyleClient:
                        styleName = ReadStyleName(payload);
                        break;
                }
            }
            if (header is not { } list)
            {
                throw new InvalidDataException("A table part has no BrtBeginList record.");
            }
            return TablePart.Create(sheetIndex, list.Id, list.Name, list.FirstRow, list.FirstColumn, list.LastRow, list.LastColumn,
                list.HeaderRowCount, list.TotalsRowCount, [.. columns], styleName);
        }

        private readonly record struct ListHeader(int Id, string Name, int FirstRow, int FirstColumn, int LastRow, int LastColumn,
            int HeaderRowCount, int TotalsRowCount);

        private static ListHeader ReadBeginList(ReadOnlySpan<byte> payload)
        {
            if (payload.Length < BeginListFixedSize
                || !Biff12.TryReadWideString(payload, BeginListFixedSize, out _, out int nameBytes)
                || !Biff12.TryReadWideString(payload, BeginListFixedSize + nameBytes, out ReadOnlySpan<char> displayName, out _))
            {
                throw Truncated("BrtBeginList");
            }
            return new ListHeader(
                Id: Biff12.ReadI32(payload, 20),
                Name: new string(displayName),
                FirstRow: Biff12.ReadI32(payload, 0),
                FirstColumn: Biff12.ReadI32(payload, 8),
                LastRow: Biff12.ReadI32(payload, 4),
                LastColumn: Biff12.ReadI32(payload, 12),
                HeaderRowCount: Biff12.ReadI32(payload, 24),
                TotalsRowCount: Biff12.ReadI32(payload, 28));
        }

        private static string ReadColumnName(ReadOnlySpan<byte> payload)
        {
            if (payload.Length < BeginListColFixedSize
                || !Biff12.TryReadWideString(payload, BeginListColFixedSize, out ReadOnlySpan<char> name, out int nameBytes)
                || !Biff12.TryReadWideString(payload, BeginListColFixedSize + nameBytes, out ReadOnlySpan<char> caption, out _))
            {
                throw Truncated("BrtBeginListCol");
            }
            if (caption.IsEmpty)
            {
                return new string(name);
            }
            return new string(caption);
        }

        private static string? ReadStyleName(ReadOnlySpan<byte> payload)
        {
            if (!Biff12.TryReadWideString(payload, StyleClientFlagsSize, out ReadOnlySpan<char> name, out _))
            {
                throw Truncated("BrtTableStyleClient");
            }
            if (name.IsEmpty)
            {
                return null;
            }
            return new string(name);
        }

        private static InvalidDataException Truncated(string record)
        {
            return new InvalidDataException($"A table part has a truncated {record} record.");
        }
    }
}
