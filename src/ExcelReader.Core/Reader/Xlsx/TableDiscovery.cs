using System.Text;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader.Xlsx
{
    internal delegate TablePart TablePartParser(int sheetIndex, ReadOnlySpan<byte> part);

    internal static class TableDiscovery
    {
        private static ReadOnlySpan<byte> TablesFolder => "xl/tables/"u8;

        private static ReadOnlySpan<byte> TableRelationshipType => "/table"u8;

        internal static List<TablePart> Load(ZipIndex zip, (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets,
            DecompressedByteCounter counter, TablePartParser parse)
        {
            List<TablePart> parts = [];
            if (!zip.HasEntryWithPrefix(TablesFolder))
            {
                return parts;
            }
            for (int i = 0; i < sheets.Length; i++)
            {
                foreach (string partPath in TablePartPaths(zip, sheets[i].Path, counter))
                {
                    using ZipPart part = zip.OpenPart(RequireEntry(zip, partPath), counter);
                    parts.Add(parse(i, part.Memory.Span));
                }
            }
            return parts;
        }

        internal static async ValueTask<List<TablePart>> LoadAsync(ZipIndex zip, (string Name, string Path, ExcelSheetVisibility Visibility)[] sheets,
            DecompressedByteCounter counter, TablePartParser parse, CancellationToken ct)
        {
            List<TablePart> parts = [];
            if (!zip.HasEntryWithPrefix(TablesFolder))
            {
                return parts;
            }
            for (int i = 0; i < sheets.Length; i++)
            {
                List<string> partPaths;
                using (ZipPart rels = await zip.OpenPartOrDefaultAsync(Utf8(RelsPath(sheets[i].Path)), counter, ct).ConfigureAwait(false))
                {
                    partPaths = ResolveTargets(sheets[i].Path, rels.Memory.Span);
                }
                foreach (string partPath in partPaths)
                {
                    using ZipPart part = await zip.OpenPartAsync(RequireEntry(zip, partPath), counter, ct).ConfigureAwait(false);
                    parts.Add(parse(i, part.Memory.Span));
                }
            }
            return parts;
        }

        internal static string RelsPath(string partPath)
        {
            int slash = partPath.LastIndexOf('/');
            return string.Concat(partPath.AsSpan(0, slash + 1), "_rels/", partPath.AsSpan(slash + 1), ".rels");
        }

        internal static string ResolvePartPath(string sourcePart, string target)
        {
            if (target.StartsWith('/'))
            {
                return target[1..];
            }
            List<string> segments = [.. sourcePart.Split('/')];
            segments.RemoveAt(segments.Count - 1);
            foreach (string segment in target.Split('/'))
            {
                ApplySegment(segments, segment);
            }
            return string.Join('/', segments);
        }

        private static void ApplySegment(List<string> segments, string segment)
        {
            if (string.Equals(segment, "..", StringComparison.Ordinal))
            {
                if (segments.Count > 0)
                {
                    segments.RemoveAt(segments.Count - 1);
                }
                return;
            }
            if (segment.Length > 0 && !string.Equals(segment, ".", StringComparison.Ordinal))
            {
                segments.Add(segment);
            }
        }

        private static List<string> TablePartPaths(ZipIndex zip, string sheetPath, DecompressedByteCounter counter)
        {
            using ZipPart rels = zip.OpenPartOrDefault(Utf8(RelsPath(sheetPath)), counter);
            return ResolveTargets(sheetPath, rels.Memory.Span);
        }

        private static List<string> ResolveTargets(string sheetPath, ReadOnlySpan<byte> rels)
        {
            List<string> targets = XlsxXml.RelationshipTargets(rels, TableRelationshipType);
            for (int i = 0; i < targets.Count; i++)
            {
                targets[i] = ResolvePartPath(sheetPath, targets[i]);
            }
            return targets;
        }

        private static ZipEntryRef RequireEntry(ZipIndex zip, string partPath)
        {
            if (!zip.TryGetEntry(Utf8(partPath), out ZipEntryRef entry))
            {
                throw new InvalidDataException($"The table part '{partPath}' is missing from the package.");
            }
            return entry;
        }

        private static byte[] Utf8(string text)
        {
            return Encoding.UTF8.GetBytes(text);
        }
    }
}
