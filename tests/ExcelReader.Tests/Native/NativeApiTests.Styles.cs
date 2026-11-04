using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Native;
using ExcelReader.Native.Typed;
using ExcelReader.Native.Writer;
using ExcelReader.Tests.Writer.Styles;

namespace ExcelReader.Tests.Native
{
    public sealed unsafe partial class NativeApiTests
    {
        private static readonly XNamespace SpreadsheetMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private static NativeCellStyleRaw UnsetStyle()
        {
            NativeBorderRaw none = new() { Style = 0, Color = -1 };
            return new NativeCellStyleRaw
            {
                StructSize = sizeof(NativeCellStyleRaw),
                FontColor = -1,
                FillColor = -1,
                Left = none,
                Right = none,
                Top = none,
                Bottom = none,
            };
        }

        private static nint OpenMemoryWriter(int format)
        {
            nint handle;
            Assert.Equal(NativeStatus.Ok, ((delegate* unmanaged<int, NativeWriteOptionsRaw*, nint*, int>)&Exports.OpenWriteHandleToMemory)(format, null, &handle));
            return handle;
        }

        private static byte[] Utf8WithTrailingJunk(string? text)
        {
            return Encoding.UTF8.GetBytes((text ?? string.Empty) + "#junk");
        }

        private static int AddStyle(nint handle, NativeCellStyleRaw style, out int styleId, string? numberFormat = null, string? fontName = null)
        {
            byte[] format = Utf8WithTrailingJunk(numberFormat);
            byte[] font = Utf8WithTrailingJunk(fontName);
            fixed (byte* formatBytes = format)
            fixed (byte* fontBytes = font)
            {
                if (numberFormat is not null)
                {
                    style.NumberFormat = formatBytes;
                    style.NumberFormatLen = Encoding.UTF8.GetByteCount(numberFormat);
                }
                if (fontName is not null)
                {
                    style.FontName = fontBytes;
                    style.FontNameLen = Encoding.UTF8.GetByteCount(fontName);
                }
                int id = -99;
                int status = ((delegate* unmanaged<nint, NativeCellStyleRaw*, int*, int>)&Exports.AddStyle)(handle, &style, &id);
                styleId = id;
                return status;
            }
        }

        private static int RegisterStyle(nint handle, NativeCellStyleRaw style, string? numberFormat = null, string? fontName = null)
        {
            Assert.Equal(NativeStatus.Ok, AddStyle(handle, style, out int styleId, numberFormat, fontName));
            return styleId;
        }

        private static int SetColumnStyle(nint handle, int column, int styleId)
        {
            return ((delegate* unmanaged<nint, int, int, int>)&Exports.SetColumnStyle)(handle, column, styleId);
        }

        private static int SetColumnWidth(nint handle, int column, double width)
        {
            return ((delegate* unmanaged<nint, int, double, int>)&Exports.SetColumnWidth)(handle, column, width);
        }

        private static int StartRowStyled(nint handle, int styleId, double height)
        {
            return ((delegate* unmanaged<nint, int, double, int>)&Exports.StartRowStyled)(handle, styleId, height);
        }

        private static int StyleNextCell(nint handle, int styleId)
        {
            return ((delegate* unmanaged<nint, int, int>)&Exports.StyleNextCell)(handle, styleId);
        }

        private static void HandleStartSheet(nint handle, string name)
        {
            byte[] utf8 = Encoding.UTF8.GetBytes(name);
            fixed (byte* bytes = utf8)
            {
                Assert.Equal(NativeStatus.Ok, ((delegate* unmanaged<nint, byte*, int, int>)&Exports.StartSheet)(handle, bytes, utf8.Length));
            }
        }

        private static void HandleWriteString(nint handle, string value)
        {
            byte[] utf8 = Encoding.UTF8.GetBytes(value);
            fixed (byte* bytes = utf8)
            {
                Assert.Equal(NativeStatus.Ok, ((delegate* unmanaged<nint, byte*, int, int>)&Exports.WriteString)(handle, bytes, utf8.Length));
            }
        }

        private static int HandleStartRow(nint handle)
        {
            return ((delegate* unmanaged<nint, int>)&Exports.StartRow)(handle);
        }

        private static int HandleWriteInt64(nint handle, long value)
        {
            return ((delegate* unmanaged<nint, long, int>)&Exports.WriteInt64)(handle, value);
        }

        private static int HandleWriteFloat64(nint handle, double value)
        {
            return ((delegate* unmanaged<nint, double, int>)&Exports.WriteFloat64)(handle, value);
        }

        private static int HandleWriteNull(nint handle, int type)
        {
            return ((delegate* unmanaged<nint, int, int>)&Exports.WriteNull)(handle, type);
        }

        private static int HandleEndRow(nint handle)
        {
            return ((delegate* unmanaged<nint, int>)&Exports.EndRow)(handle);
        }

        private static int HandleEndSheet(nint handle)
        {
            return ((delegate* unmanaged<nint, int>)&Exports.EndSheet)(handle);
        }

        private static int HandleClose(nint handle)
        {
            return ((delegate* unmanaged<nint, int>)&Exports.CloseWriteHandle)(handle);
        }

        private static int HandleAbort(nint handle)
        {
            return ((delegate* unmanaged<nint, int>)&Exports.AbortWriteHandle)(handle);
        }

        private static void WriteOneRow(nint handle, long value)
        {
            Assert.Equal(NativeStatus.Ok, HandleStartRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, value));
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
        }

        private static byte[] CloseAndReadBytes(nint handle)
        {
            NativeBuffer buffer;
            Assert.Equal(NativeStatus.Ok, ((delegate* unmanaged<nint, NativeBuffer*, int>)&Exports.WriteHandleBytes)(handle, &buffer));
            byte[] bytes = new ReadOnlySpan<byte>((void*)buffer.Data, checked((int)buffer.Length)).ToArray();
            ((delegate* unmanaged<NativeBuffer*, void>)&Exports.FreeBuffer)(&buffer);
            Assert.Equal(NativeStatus.Ok, HandleClose(handle));
            return bytes;
        }

        private static XElement XlsxSheetXml(byte[] file)
        {
            using var zip = new ZipArchive(new MemoryStream(file), ZipArchiveMode.Read);
            return XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open()).Root!;
        }

        [Fact]
        public void AddStyleDedupesAndZeroIsGeneral()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            NativeCellStyleRaw bold = UnsetStyle() with { FontFlags = 1 };

            int first = RegisterStyle(handle, bold);
            int second = RegisterStyle(handle, bold);
            int italic = RegisterStyle(handle, UnsetStyle() with { FontFlags = 2 });

            Assert.True(first > 0);
            Assert.Equal(first, second);
            Assert.NotEqual(first, italic);
            HandleStartSheet(handle, "S");
            Assert.Equal(NativeStatus.Ok, SetColumnStyle(handle, 0, 0));
            WriteOneRow(handle, 1);
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            byte[] bytes = CloseAndReadBytes(handle);

            ProbedStyle cell = StyleProbe.Xlsx(bytes, 0, 0);
            Assert.False(cell.Bold);
            Assert.False(cell.Italic);
            Assert.Null(cell.NumberFormat);
        }

        [Fact]
        public void ColumnStyleAndWidth()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            int money = RegisterStyle(handle, UnsetStyle() with { FillColor = 0xFF0000 }, numberFormat: "#,##0.00 €");
            HandleStartSheet(handle, "S");

            Assert.Equal(NativeStatus.Ok, SetColumnStyle(handle, 1, money));
            Assert.Equal(NativeStatus.Ok, SetColumnWidth(handle, 1, 20));
            Assert.Equal(NativeStatus.Ok, HandleStartRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 1));
            Assert.Equal(NativeStatus.Ok, HandleWriteFloat64(handle, 2.5));
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            byte[] bytes = CloseAndReadBytes(handle);

            Assert.Null(StyleProbe.Xlsx(bytes, 0, 0).FillRgb);
            ProbedStyle styled = StyleProbe.Xlsx(bytes, 0, 1);
            Assert.Equal(0xFF0000, styled.FillRgb);
            Assert.Equal("#,##0.00 €", styled.NumberFormat);
            XElement col = Assert.Single(XlsxSheetXml(bytes).Descendants(SpreadsheetMain + "col"));
            Assert.Equal("2", col.Attribute("min")!.Value);
            Assert.Equal(20, double.Parse(col.Attribute("width")!.Value, CultureInfo.InvariantCulture));
        }

        [Fact]
        public void RowStyleAndHeight()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            int red = RegisterStyle(handle, UnsetStyle() with { FillColor = 0xFF0000 });
            int bold = RegisterStyle(handle, UnsetStyle() with { FontFlags = 1 });
            HandleStartSheet(handle, "S");
            Assert.Equal(NativeStatus.Ok, SetColumnStyle(handle, 0, red));

            Assert.Equal(NativeStatus.Ok, StartRowStyled(handle, bold, 30));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 1));
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
            Assert.Equal(NativeStatus.Ok, StartRowStyled(handle, 0, 0));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 2));
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            byte[] bytes = CloseAndReadBytes(handle);

            ProbedStyle styledRow = StyleProbe.Xlsx(bytes, 0, 0);
            Assert.True(styledRow.Bold);
            Assert.Null(styledRow.FillRgb);
            Assert.Equal(30, StyleProbe.RowHeightXlsx(bytes, 0));
            Assert.Equal(0xFF0000, StyleProbe.Xlsx(bytes, 1, 0).FillRgb);
            Assert.Null(StyleProbe.RowHeightXlsx(bytes, 1));
        }

        [Fact]
        public void StyleNextCellAppliesOnce()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            int bold = RegisterStyle(handle, UnsetStyle() with { FontFlags = 1 });
            HandleStartSheet(handle, "S");

            Assert.Equal(NativeStatus.Ok, HandleStartRow(handle));
            Assert.Equal(NativeStatus.Ok, StyleNextCell(handle, bold));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 1));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 2));
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            byte[] bytes = CloseAndReadBytes(handle);

            Assert.True(StyleProbe.Xlsx(bytes, 0, 0).Bold);
            Assert.False(StyleProbe.Xlsx(bytes, 0, 1).Bold);
        }

        [Fact]
        public void StyledNullIsStyledBlank()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            int yellow = RegisterStyle(handle, UnsetStyle() with { FillColor = 0xFFFF00 });
            HandleStartSheet(handle, "S");

            Assert.Equal(NativeStatus.Ok, HandleStartRow(handle));
            Assert.Equal(NativeStatus.Ok, StyleNextCell(handle, yellow));
            Assert.Equal(NativeStatus.Ok, HandleWriteNull(handle, NativeColumnType.String));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 2));
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            byte[] bytes = CloseAndReadBytes(handle);

            XElement blank = XlsxSheetXml(bytes).Descendants(SpreadsheetMain + "c").First(c => !c.HasElements);
            Assert.Equal(0xFFFF00, StyleProbe.XlsxAt(bytes, int.Parse(blank.Attribute("s")!.Value, CultureInfo.InvariantCulture)).FillRgb);
            Assert.Null(StyleProbe.Xlsx(bytes, 0, 1).FillRgb);
        }

        [Fact]
        public void PendingStyleAtEndRowFails()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            int bold = RegisterStyle(handle, UnsetStyle() with { FontFlags = 1 });
            HandleStartSheet(handle, "S");
            Assert.Equal(NativeStatus.Ok, HandleStartRow(handle));
            Assert.Equal(NativeStatus.Ok, StyleNextCell(handle, bold));

            Assert.Equal(NativeStatus.Error, HandleEndRow(handle));
            Assert.NotEmpty(NativeApi.LastErrorText());
            Assert.Equal(NativeStatus.Ok, HandleAbort(handle));
        }

        [Fact]
        public void PendingStyleAtEndRowLeavesTheHandleWritable()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            int bold = RegisterStyle(handle, UnsetStyle() with { FontFlags = 1 });
            HandleStartSheet(handle, "S");
            Assert.Equal(NativeStatus.Ok, HandleStartRow(handle));
            Assert.Equal(NativeStatus.Ok, StyleNextCell(handle, bold));
            Assert.Equal(NativeStatus.Error, HandleEndRow(handle));

            WriteOneRow(handle, 2);
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            byte[] bytes = CloseAndReadBytes(handle);

            using IExcelWorkbook workbook = Excel.Open(bytes.AsMemory());
            Assert.Equal(["", "2"], ReadCells(workbook.SheetAt(0)), StringComparer.Ordinal);
        }

        [Fact]
        public void InvalidStyleDoesNotPoison()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            NativeCellStyleRaw badColor = UnsetStyle() with { FontColor = 0x1000000 };
            NativeCellStyleRaw badBorder = UnsetStyle() with { Left = new NativeBorderRaw { Style = 99, Color = -1 } };

            Assert.Equal(NativeStatus.InvalidArgument, AddStyle(handle, badColor, out int badId));
            Assert.Contains("font_color", NativeApi.LastErrorText(), StringComparison.Ordinal);
            Assert.Equal(0, badId);
            Assert.Equal(NativeStatus.InvalidArgument, AddStyle(handle, badBorder, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());

            int bold = RegisterStyle(handle, UnsetStyle() with { FontFlags = 1 });
            HandleStartSheet(handle, "S");
            Assert.Equal(NativeStatus.Ok, StartRowStyled(handle, bold, 0));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 1));
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            Assert.Equal(NativeStatus.Ok, HandleClose(handle));
        }

        [Theory]
        [InlineData("struct_size")]
        [InlineData("font_flags")]
        [InlineData("wrap_text")]
        [InlineData("fill_color")]
        [InlineData("right")]
        [InlineData("number_format_len")]
        [InlineData("font_name_len")]
        [InlineData("font_size")]
        [InlineData("horizontal_alignment")]
        [InlineData("indent")]
        public void EachInvalidStyleFieldIsRejected(string field)
        {
            NativeCellStyleRaw style = UnsetStyle();
            switch (field)
            {
                case "struct_size":
                    style.StructSize = 100;
                    break;
                case "font_flags":
                    style.FontFlags = 16;
                    break;
                case "wrap_text":
                    style.WrapText = 2;
                    break;
                case "fill_color":
                    style.FillColor = -2;
                    break;
                case "right":
                    style.Right = new NativeBorderRaw { Style = 1, Color = 0x1000000 };
                    break;
                case "number_format_len":
                    style.NumberFormat = (byte*)1;
                    style.NumberFormatLen = -1;
                    break;
                case "font_name_len":
                    style.FontName = (byte*)1;
                    style.FontNameLen = -1;
                    break;
                case "font_size":
                    style.FontSize = 500;
                    break;
                case "horizontal_alignment":
                    style.HorizontalAlignment = 8;
                    break;
                default:
                    style.Indent = 16;
                    break;
            }
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);

            Assert.Equal(NativeStatus.InvalidArgument, AddStyle(handle, style, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
            Assert.Equal(NativeStatus.Ok, HandleAbort(handle));
        }

        [Fact]
        public void StyleCallsRejectNullPointersAndUnknownHandles()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            NativeCellStyleRaw style = UnsetStyle();
            int id;
            delegate* unmanaged<nint, NativeCellStyleRaw*, int*, int> addStyle = &Exports.AddStyle;

            Assert.Equal(NativeStatus.InvalidArgument, addStyle(handle, null, &id));
            Assert.Equal(NativeStatus.InvalidArgument, addStyle(handle, &style, null));
            Assert.Equal(NativeStatus.InvalidHandle, addStyle(0, &style, &id));
            Assert.Equal(NativeStatus.InvalidHandle, SetColumnStyle(0, 0, 0));
            Assert.Equal(NativeStatus.InvalidHandle, SetColumnWidth(0, 0, 10));
            Assert.Equal(NativeStatus.InvalidHandle, StartRowStyled(0, 0, 0));
            Assert.Equal(NativeStatus.InvalidHandle, StyleNextCell(0, 0));
            Assert.Equal(NativeStatus.Ok, HandleAbort(handle));
        }

        [Fact]
        public void ColumnStyleAfterFirstRowFails()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
            Assert.Equal(NativeStatus.Error, SetColumnStyle(handle, 0, 0));
            HandleStartSheet(handle, "S");

            Assert.Equal(NativeStatus.InvalidArgument, SetColumnStyle(handle, -1, 0));
            Assert.Equal(NativeStatus.InvalidArgument, SetColumnStyle(handle, 0, 999));
            Assert.Equal(NativeStatus.InvalidArgument, SetColumnWidth(handle, 0, -1));
            Assert.Equal(NativeStatus.InvalidArgument, StartRowStyled(handle, 999, 0));
            Assert.Equal(NativeStatus.InvalidArgument, StartRowStyled(handle, 0, 500));
            Assert.Equal(NativeStatus.Ok, HandleStartRow(handle));
            Assert.Equal(NativeStatus.InvalidArgument, StyleNextCell(handle, 999));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 1));
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));

            Assert.Equal(NativeStatus.Error, SetColumnStyle(handle, 0, 0));
            Assert.NotEmpty(NativeApi.LastErrorText());
            Assert.Equal(NativeStatus.Error, SetColumnWidth(handle, 0, 12));
            Assert.Equal(NativeStatus.Error, StyleNextCell(handle, 0));
            WriteOneRow(handle, 2);
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            byte[] bytes = CloseAndReadBytes(handle);

            using IExcelWorkbook workbook = Excel.Open(bytes.AsMemory());
            Assert.Equal(["1", "2"], ReadCells(workbook.SheetAt(0)), StringComparer.Ordinal);
        }

        [Fact]
        public void XlsColorLimitIsError()
        {
            nint handle = OpenMemoryWriter(NativeFormat.Xls);
            for (int color = 0; color < 56; color++)
            {
                RegisterStyle(handle, UnsetStyle() with { FillColor = color });
            }

            Assert.Equal(NativeStatus.Error, AddStyle(handle, UnsetStyle() with { FillColor = 56 }, out _));
            Assert.NotEmpty(NativeApi.LastErrorText());
            HandleStartSheet(handle, "S");
            WriteOneRow(handle, 1);
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            Assert.Equal(NativeStatus.Ok, HandleClose(handle));
        }

        [Fact]
        public void StylesOnCsvAreNoOps()
        {
            nint plain = OpenMemoryWriter(NativeFormat.Csv);
            HandleStartSheet(plain, "S");
            Assert.Equal(NativeStatus.Ok, HandleStartRow(plain));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(plain, 1));
            HandleWriteString(plain, "a");
            Assert.Equal(NativeStatus.Ok, HandleEndRow(plain));
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(plain));
            byte[] expected = CloseAndReadBytes(plain);

            nint styled = OpenMemoryWriter(NativeFormat.Csv);
            int bold = RegisterStyle(styled, UnsetStyle() with { FontFlags = 1, FillColor = 0x00FF00 }, numberFormat: "0.00");
            HandleStartSheet(styled, "S");
            Assert.Equal(NativeStatus.Ok, SetColumnStyle(styled, 0, bold));
            Assert.Equal(NativeStatus.Ok, SetColumnWidth(styled, 0, 20));
            Assert.Equal(NativeStatus.Ok, StartRowStyled(styled, bold, 30));
            Assert.Equal(NativeStatus.Ok, StyleNextCell(styled, bold));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(styled, 1));
            HandleWriteString(styled, "a");
            Assert.Equal(NativeStatus.Ok, HandleEndRow(styled));
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(styled));

            Assert.Equal(expected, CloseAndReadBytes(styled));
        }

        [Theory]
        [InlineData(NativeFormat.Xlsb)]
        [InlineData(NativeFormat.Xls)]
        public void XlsbAndXlsWriteAndReopen(int format)
        {
            nint handle = OpenMemoryWriter(format);
            NativeCellStyleRaw column = UnsetStyle() with
            {
                FontFlags = 1,
                FillColor = 0x00FF00,
                HorizontalAlignment = 2,
                Left = new NativeBorderRaw { Style = 1, Color = 0xFF0000 },
            };
            int columnStyle = RegisterStyle(handle, column, numberFormat: "0.00", fontName: "Courier New");
            int italic = RegisterStyle(handle, UnsetStyle() with { FontFlags = 2 });
            HandleStartSheet(handle, "S");
            Assert.Equal(NativeStatus.Ok, SetColumnStyle(handle, 0, columnStyle));
            Assert.Equal(NativeStatus.Ok, SetColumnWidth(handle, 0, 18));

            Assert.Equal(NativeStatus.Ok, StartRowStyled(handle, italic, 25));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 1));
            HandleWriteString(handle, "x");
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleStartRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleWriteInt64(handle, 2));
            Assert.Equal(NativeStatus.Ok, StyleNextCell(handle, italic));
            HandleWriteString(handle, "y");
            Assert.Equal(NativeStatus.Ok, HandleEndRow(handle));
            Assert.Equal(NativeStatus.Ok, HandleEndSheet(handle));
            byte[] bytes = CloseAndReadBytes(handle);

            using (IExcelWorkbook workbook = Excel.Open(bytes.AsMemory()))
            {
                Assert.Equal(["1|x", "2|y"], ReadCells(workbook.SheetAt(0)), StringComparer.Ordinal);
            }
            bool xlsb = format == NativeFormat.Xlsb;
            ProbedStyle styledColumn = xlsb ? StyleProbe.Xlsb(bytes, 1, 0) : StyleProbe.Xls(bytes, 1, 0);
            Assert.True(styledColumn.Bold);
            Assert.Equal(0x00FF00, styledColumn.FillRgb);
            Assert.Equal("0.00", styledColumn.NumberFormat);
            Assert.Equal("Courier New", styledColumn.FontName);
            Assert.Equal(ExcelHorizontalAlignment.Center, styledColumn.Horizontal);
            Assert.Equal((ExcelBorderStyle.Thin, (int?)0xFF0000), styledColumn.Left);
            ProbedStyle styledRow = xlsb ? StyleProbe.Xlsb(bytes, 0, 0) : StyleProbe.Xls(bytes, 0, 0);
            Assert.True(styledRow.Italic);
            Assert.False(styledRow.Bold);
            Assert.True((xlsb ? StyleProbe.Xlsb(bytes, 1, 1) : StyleProbe.Xls(bytes, 1, 1)).Italic);
            Assert.Equal(25, xlsb ? StyleProbe.RowHeightXlsb(bytes, 0) : StyleProbe.RowHeightXls(bytes, 0));
        }

        private static nint OpenParallelMemoryWriter(int format, int degreeOfParallelism, int batchRows)
        {
            NativeWriteOptionsRaw raw = DefaultWriteOptionsRaw();
            raw.DegreeOfParallelism = degreeOfParallelism;
            raw.BatchRows = batchRows;
            nint handle;
            Assert.Equal(NativeStatus.Ok, ((delegate* unmanaged<int, NativeWriteOptionsRaw*, nint*, int>)&Exports.OpenWriteHandleToMemory)(format, &raw, &handle));
            return handle;
        }

        private static IntPtr Unmanaged(List<IntPtr> blocks, int size)
        {
            IntPtr block = Marshal.AllocHGlobal(Math.Max(size, 1));
            blocks.Add(block);
            return block;
        }

        private static NativeColumnSpecRaw* RawSpecs(NativeColumnSpec[] specs, List<IntPtr> blocks)
        {
            var raw = (NativeColumnSpecRaw*)Unmanaged(blocks, specs.Length * sizeof(NativeColumnSpecRaw));
            for (int index = 0; index < specs.Length; index++)
            {
                string[] names = specs[index].Names;
                var pointers = (byte**)Unmanaged(blocks, names.Length * sizeof(nint));
                var lengths = (int*)Unmanaged(blocks, names.Length * sizeof(int));
                for (int n = 0; n < names.Length; n++)
                {
                    byte[] utf8 = Encoding.UTF8.GetBytes(names[n]);
                    pointers[n] = (byte*)Unmanaged(blocks, utf8.Length);
                    utf8.CopyTo(new Span<byte>(pointers[n], utf8.Length));
                    lengths[n] = utf8.Length;
                }
                raw[index] = new NativeColumnSpecRaw { Names = pointers, NameLens = lengths, NameCount = names.Length, Index = index, Type = specs[index].Type };
            }
            return raw;
        }

        private static NativeCellStyleRaw* RawStyle(NativeCellStyleRaw? style, List<IntPtr> blocks)
        {
            if (style is not { } value)
            {
                return null;
            }
            var copy = (NativeCellStyleRaw*)Unmanaged(blocks, sizeof(NativeCellStyleRaw));
            *copy = value;
            return copy;
        }

        // Builds xl_sheet_options in a 64-byte block pre-filled with 0xFF, so a smaller struct_size proves the
        // bytes past it are never read.
        private static NativeSheetOptionsRaw* RawSheetOptions(SheetStyleArgs args, List<IntPtr> blocks)
        {
            NativeSheetOptionsRaw options = new() { StructSize = args.StructSize };
            if (args.TableName is not null)
            {
                byte[] utf8 = Encoding.UTF8.GetBytes(args.TableName);
                options.TableName = (byte*)Unmanaged(blocks, utf8.Length);
                utf8.CopyTo(new Span<byte>(options.TableName, utf8.Length));
                options.TableNameLen = utf8.Length;
            }
            if (args.ColumnStyles is not null)
            {
                options.ColumnStyles = (NativeCellStyleRaw**)Unmanaged(blocks, args.ColumnStyles.Length * sizeof(nint));
                for (int index = 0; index < args.ColumnStyles.Length; index++)
                {
                    options.ColumnStyles[index] = RawStyle(args.ColumnStyles[index], blocks);
                }
            }
            if (args.Widths is not null)
            {
                options.ColumnWidths = (double*)Unmanaged(blocks, args.Widths.Length * sizeof(double));
                args.Widths.CopyTo(new Span<double>(options.ColumnWidths, args.Widths.Length));
            }
            options.HeaderStyle = RawStyle(args.HeaderStyle, blocks);
            var block = (byte*)Unmanaged(blocks, sizeof(NativeSheetOptionsRaw));
            new Span<byte>(block, sizeof(NativeSheetOptionsRaw)).Fill(0xFF);
            new ReadOnlySpan<byte>(&options, args.StructSize).CopyTo(new Span<byte>(block, args.StructSize));
            return (NativeSheetOptionsRaw*)block;
        }

        private sealed record SheetStyleArgs(
            NativeCellStyleRaw?[]? ColumnStyles = null, double[]? Widths = null, NativeCellStyleRaw? HeaderStyle = null,
            string? TableName = null, int StructSize = 64);

        private static int WriteTypedSheet(nint handle, string name, NativeColumnSpec[] specs, NativeTable table, SheetStyleArgs args)
        {
            List<IntPtr> blocks = [];
            try
            {
                NativeColumnSpecRaw* rawSpecs = RawSpecs(specs, blocks);
                NativeSheetOptionsRaw* options = RawSheetOptions(args, blocks);
                byte[] utf8 = Encoding.UTF8.GetBytes(name);
                fixed (byte* nameBytes = utf8)
                {
                    return ((delegate* unmanaged<nint, byte*, int, NativeColumnSpecRaw*, NativeTable*, NativeSheetOptionsRaw*, int>)&Exports.WriteSheet)(
                        handle, nameBytes, utf8.Length, rawSpecs, &table, options);
                }
            }
            finally
            {
                foreach (IntPtr block in blocks)
                {
                    Marshal.FreeHGlobal(block);
                }
            }
        }

        private static byte[] WriteOneTypedSheet(int format, NativeColumnSpec[] specs, NativeTable table, SheetStyleArgs args)
        {
            nint handle = OpenMemoryWriter(format);
            Assert.Equal(NativeStatus.Ok, WriteTypedSheet(handle, "S", specs, table, args));
            return CloseAndReadBytes(handle);
        }

        private static ProbedStyle Probe(int format, byte[] bytes, int row, int column)
        {
            switch (format)
            {
                case NativeFormat.Xlsb:
                    return StyleProbe.Xlsb(bytes, row, column);
                case NativeFormat.Xls:
                    return StyleProbe.Xls(bytes, row, column);
                default:
                    return StyleProbe.Xlsx(bytes, row, column);
            }
        }

        private static string ContentDigest(byte[] file)
        {
            if (file is not [(byte)'P', (byte)'K', ..])
            {
                return Convert.ToHexString(SHA256.HashData(file));
            }
            using IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            using var zip = new ZipArchive(new MemoryStream(file), ZipArchiveMode.Read);
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                hash.AppendData(Encoding.UTF8.GetBytes(entry.FullName));
                using var content = new MemoryStream();
                entry.Open().CopyTo(content);
                hash.AppendData(content.ToArray());
            }
            return Convert.ToHexString(hash.GetHashAndReset());
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, false)]
        [InlineData(NativeFormat.Xlsx, true)]
        [InlineData(NativeFormat.Xlsb, false)]
        [InlineData(NativeFormat.Xlsb, true)]
        [InlineData(NativeFormat.Xls, false)]
        public void TypedColumnStylesDataOnly(int format, bool asTable)
        {
            NativeTable sales = BuildQtyNameTable([3L, 7L], ["widget", "gadget"]);
            try
            {
                NativeCellStyleRaw bold = UnsetStyle() with { FontFlags = NativeCellStyle.Bold };
                byte[] bytes = WriteOneTypedSheet(format, QtyNameSpecs(), sales,
                    new SheetStyleArgs(ColumnStyles: [bold, bold], TableName: asTable ? "Sales" : null));

                for (int column = 0; column < 2; column++)
                {
                    ProbedStyle header = Probe(format, bytes, 0, column);
                    Assert.False(header.Bold);
                    Assert.Null(header.FillRgb);
                    Assert.True(Probe(format, bytes, 1, column).Bold);
                    Assert.True(Probe(format, bytes, 2, column).Bold);
                }
                using IExcelWorkbook workbook = Excel.Open(bytes.AsMemory());
                Assert.Equal(["qty|name", "3|widget", "7|gadget"], ReadCells(workbook.SheetAt(0)), StringComparer.Ordinal);
                Assert.Equal(asTable ? 1 : 0, workbook.Tables.Count);
            }
            finally
            {
                FreeBuiltTable(ref sales);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void TypedHeaderStyleAndWidths(bool asTable)
        {
            NativeTable sales = BuildQtyNameTable([3L, 7L], ["widget", "gadget"]);
            try
            {
                NativeCellStyleRaw header = UnsetStyle() with { FontFlags = NativeCellStyle.Italic, FillColor = 0xDDEEFF };
                byte[] bytes = WriteOneTypedSheet(NativeFormat.Xlsx, QtyNameSpecs(), sales,
                    new SheetStyleArgs(Widths: [0, 24.5], HeaderStyle: header, TableName: asTable ? "Sales" : null));

                for (int column = 0; column < 2; column++)
                {
                    ProbedStyle headerCell = StyleProbe.Xlsx(bytes, 0, column);
                    Assert.True(headerCell.Italic);
                    Assert.Equal(0xDDEEFF, headerCell.FillRgb);
                    ProbedStyle data = StyleProbe.Xlsx(bytes, 1, column);
                    Assert.False(data.Italic);
                    Assert.Null(data.FillRgb);
                }
                XElement col = Assert.Single(XlsxSheetXml(bytes).Descendants(SpreadsheetMain + "col"));
                Assert.Equal("2", col.Attribute("min")!.Value);
                Assert.Equal(24.5, double.Parse(col.Attribute("width")!.Value, CultureInfo.InvariantCulture));
            }
            finally
            {
                FreeBuiltTable(ref sales);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, true)]
        [InlineData(NativeFormat.Xlsx, false)]
        [InlineData(NativeFormat.Xlsb, true)]
        [InlineData(NativeFormat.Xlsb, false)]
        public void TypedStyleKeepsTemporalFormat(int format, bool hasHeader)
        {
            NativeTable temporal = BuildTemporalTable(19000, 3_600_000_000L, 1_700_000_000_000_000L);
            try
            {
                NativeColumnSpec[] specs =
                [
                    new() { Names = hasHeader ? ["date"] : [], Type = NativeColumnType.Date },
                    new() { Names = hasHeader ? ["time"] : [], Type = NativeColumnType.Time },
                    new() { Names = hasHeader ? ["stamp"] : [], Type = NativeColumnType.Timestamp },
                ];
                NativeCellStyleRaw bold = UnsetStyle() with { FontFlags = NativeCellStyle.Bold };
                byte[] bytes = WriteOneTypedSheet(format, specs, temporal, new SheetStyleArgs(ColumnStyles: [bold, bold, bold]));

                int row = hasHeader ? 1 : 0;
                ProbedStyle time = Probe(format, bytes, row, 1);
                Assert.True(time.Bold);
                Assert.Equal("hh:mm:ss", time.NumberFormat);
                ProbedStyle stamp = Probe(format, bytes, row, 2);
                Assert.True(stamp.Bold);
                Assert.Equal("yyyy-mm-dd hh:mm:ss", stamp.NumberFormat);
                Assert.True(Probe(format, bytes, row, 0).Bold);
                using IExcelWorkbook workbook = Excel.Open(bytes.AsMemory());
                using IExcelRowEnumerator rows = workbook.SheetAt(0).GetEnumerator();
                for (int skipped = 0; skipped <= row; skipped++)
                {
                    Assert.True(rows.MoveNext());
                }
                Assert.Equal([CellType.Date, CellType.Date, CellType.Date], [rows.Current[0].Type, rows.Current[1].Type, rows.Current[2].Type]);
            }
            finally
            {
                FreeBuiltTable(ref temporal);
            }
        }

        [Fact]
        public void TypedStyleExplicitNumberFormatWins()
        {
            NativeTable temporal = BuildTemporalTable(19000, 3_600_000_000L, 1_700_000_000_000_000L);
            List<IntPtr> blocks = [];
            try
            {
                byte[] format = Encoding.UTF8.GetBytes("[h]:mm");
                var formatBytes = (byte*)Unmanaged(blocks, format.Length);
                format.CopyTo(new Span<byte>(formatBytes, format.Length));
                NativeCellStyleRaw duration = UnsetStyle() with { NumberFormat = formatBytes, NumberFormatLen = format.Length };
                NativeColumnSpec[] specs =
                [
                    new() { Names = ["date"], Type = NativeColumnType.Date },
                    new() { Names = ["time"], Type = NativeColumnType.Time },
                    new() { Names = ["stamp"], Type = NativeColumnType.Timestamp },
                ];
                byte[] bytes = WriteOneTypedSheet(NativeFormat.Xlsx, specs, temporal, new SheetStyleArgs(ColumnStyles: [null, duration, null]));

                Assert.Equal("[h]:mm", StyleProbe.Xlsx(bytes, 1, 1).NumberFormat);
                Assert.Equal("yyyy-mm-dd hh:mm:ss", StyleProbe.Xlsx(bytes, 1, 2).NumberFormat);
            }
            finally
            {
                foreach (IntPtr block in blocks)
                {
                    Marshal.FreeHGlobal(block);
                }
                FreeBuiltTable(ref temporal);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, false)]
        [InlineData(NativeFormat.Xlsx, true)]
        [InlineData(NativeFormat.Xlsb, false)]
        [InlineData(NativeFormat.Xlsb, true)]
        public void TypedStylesInParallelWrite(int format, bool asTable)
        {
            const int RowCount = 9;
            NativeTable table = BuildEveryTypeTable(RowCount);
            try
            {
                NativeCellStyleRaw bold = UnsetStyle() with { FontFlags = NativeCellStyle.Bold };
                NativeCellStyleRaw header = UnsetStyle() with { FontFlags = NativeCellStyle.Italic };
                nint handle = OpenParallelMemoryWriter(format, degreeOfParallelism: 4, batchRows: 2);
                Assert.Equal(NativeStatus.Ok, WriteTypedSheet(handle, "S", EveryTypeSpecs(), table,
                    new SheetStyleArgs(ColumnStyles: [null, bold, null, null, null, bold, null], HeaderStyle: header, TableName: asTable ? "Data" : null)));
                byte[] bytes = CloseAndReadBytes(handle);

                ProbedStyle headerCell = Probe(format, bytes, 0, 1);
                Assert.True(headerCell.Italic);
                Assert.False(headerCell.Bold);
                foreach (int row in new[] { 2, 4, 5, 6, 7, 8 })
                {
                    Assert.True(Probe(format, bytes, row, 1).Bold, $"row {row}");
                    Assert.False(Probe(format, bytes, row, 2).Bold);
                    ProbedStyle time = Probe(format, bytes, row, 5);
                    Assert.True(time.Bold);
                    Assert.Equal("hh:mm:ss", time.NumberFormat);
                }
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, "D397F2A61315BC0EACAAE77AED1D92C9DBEF4F3D8F0EEACD69B2E1271C486EC6")]
        [InlineData(NativeFormat.Xlsb, "45249BA1281F1805EAECD53F29D817310ECFED9314BD39CB24D612F6DDDE44FD")]
        [InlineData(NativeFormat.Xls, "C343CC4143F315F5C9FD32CF0F0AAFF51CA1C990F828952EF65734EE9DEA393E")]
        [InlineData(NativeFormat.Csv, "FE3FE9C07148018D946635635108234CE122A47A88AEBEC5FDC6D73F980AD7E0")]
        public void WriteTypedStillWorks(int format, string expectedDigest)
        {
            NativeTable table = BuildEveryTypeTable(7);
            try
            {
                Assert.Equal(NativeStatus.Ok, WriteApi.WriteTypedToMemory(format, EveryTypeSpecs(), table, DefaultWriteOptions(), out byte[]? bytes));
                Assert.Equal(expectedDigest, ContentDigest(bytes!));
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Theory]
        [InlineData(NativeFormat.Xlsx, "299C298F0C874FC9FB81060C72D6F335CD4F8EB72E5593C7031B11BE2703E353")]
        [InlineData(NativeFormat.Xlsb, "3CA9DA1D60A9A5C2C6AE5DA25E302535C072BDF18C829E3432FDB9113730B410")]
        public void Abi6SizesStillWork(int format, string expectedDigest)
        {
            NativeTable table = BuildEveryTypeTable(7);
            try
            {
                byte[] bytes = WriteOneTypedSheet(format, EveryTypeSpecs(), table, new SheetStyleArgs(TableName: "Data", StructSize: NativeSheetOptionsRaw.V6Size));
                Assert.Equal(expectedDigest, ContentDigest(bytes));
            }
            finally
            {
                FreeBuiltTable(ref table);
            }
        }

        [Fact]
        public void ColumnStylesInvalidEntry()
        {
            NativeTable sales = BuildQtyNameTable([3L], ["widget"]);
            try
            {
                NativeCellStyleRaw bold = UnsetStyle() with { FontFlags = NativeCellStyle.Bold };
                NativeCellStyleRaw badColor = UnsetStyle() with { FontColor = 0x1000000 };
                nint handle = OpenMemoryWriter(NativeFormat.Xlsx);
                SheetStyleArgs[] invalid =
                [
                    new(ColumnStyles: [bold, badColor]),
                    new(HeaderStyle: badColor),
                    new(Widths: [10, -1]),
                    new(Widths: [256, 0]),
                    new(Widths: [double.NaN, 0]),
                    new(StructSize: 48),
                ];
                foreach (SheetStyleArgs args in invalid)
                {
                    Assert.Equal(NativeStatus.InvalidArgument, WriteTypedSheet(handle, "S", QtyNameSpecs(), sales, args));
                    Assert.NotEmpty(NativeApi.LastErrorText());
                }
                Assert.Contains("column_styles[1]", WriteTypedSheetError(handle, sales, invalid[0]), StringComparison.Ordinal);

                Assert.Equal(NativeStatus.Ok, WriteTypedSheet(handle, "S", QtyNameSpecs(), sales, new SheetStyleArgs(ColumnStyles: [bold, null])));
                byte[] bytes = CloseAndReadBytes(handle);
                using IExcelWorkbook workbook = Excel.Open(bytes.AsMemory());
                Assert.Equal(1, workbook.SheetCount);
                Assert.True(StyleProbe.Xlsx(bytes, 1, 0).Bold);
                Assert.False(StyleProbe.Xlsx(bytes, 1, 1).Bold);
            }
            finally
            {
                FreeBuiltTable(ref sales);
            }
        }

        private static string WriteTypedSheetError(nint handle, NativeTable table, SheetStyleArgs args)
        {
            Assert.Equal(NativeStatus.InvalidArgument, WriteTypedSheet(handle, "S", QtyNameSpecs(), table, args));
            return NativeApi.LastErrorText();
        }

        [Fact]
        public void ColumnStyleLimitIsErrorWithoutPoisoning()
        {
            NativeTable sales = BuildQtyNameTable([3L], ["widget"]);
            try
            {
                nint handle = OpenMemoryWriter(NativeFormat.Xls);
                for (int color = 0; color < 56; color++)
                {
                    RegisterStyle(handle, UnsetStyle() with { FillColor = color });
                }

                Assert.Equal(NativeStatus.Error, WriteTypedSheet(handle, "S", QtyNameSpecs(), sales,
                    new SheetStyleArgs(ColumnStyles: [UnsetStyle() with { FillColor = 56 }, null])));
                Assert.NotEmpty(NativeApi.LastErrorText());
                Assert.Equal(NativeStatus.Ok, WriteTypedSheet(handle, "S", QtyNameSpecs(), sales, new SheetStyleArgs()));
                byte[] bytes = CloseAndReadBytes(handle);
                using IExcelWorkbook workbook = Excel.Open(bytes.AsMemory());
                Assert.Equal(1, workbook.SheetCount);
            }
            finally
            {
                FreeBuiltTable(ref sales);
            }
        }

        [Fact]
        public void TypedStylesOnCsvAreNoOps()
        {
            NativeTable sales = BuildQtyNameTable([3L, 7L], ["widget", "gadget"]);
            try
            {
                NativeCellStyleRaw bold = UnsetStyle() with { FontFlags = NativeCellStyle.Bold };
                byte[] plain = WriteOneTypedSheet(NativeFormat.Csv, QtyNameSpecs(), sales, new SheetStyleArgs());
                byte[] styled = WriteOneTypedSheet(NativeFormat.Csv, QtyNameSpecs(), sales,
                    new SheetStyleArgs(ColumnStyles: [bold, bold], Widths: [10, 20], HeaderStyle: bold));

                Assert.Equal(plain, styled);
            }
            finally
            {
                FreeBuiltTable(ref sales);
            }
        }
    }
}
