namespace ExcelReader.Core.Writer.Internal
{
    internal readonly record struct BoundLayout<TRow, T>(string[] Headers, Action<TRow, T> WriteRow, int HeaderStyleId, int[]? ColumnStyleIds)
        where TRow : IRowWriter;

    internal static class LayoutBinding
    {
        /// <summary>Resolves a layout's styles to the sheet's workbook ids. Registers every style before any row is written.</summary>
        internal static BoundLayout<TRow, T> Bind<TRow, T>(ISheetWriter<TRow> sheet, ExcelRecordLayout<T> layout)
            where TRow : IRowWriter
        {
            (string[] headers, Action<TRow, T> writeRow, CellStyle?[]? columnStyles) = layout.Columns<TRow>();
            CellStyle? headerStyle = layout.HeaderStyle;
            if (columnStyles is null && headerStyle is null)
            {
                return new(headers, writeRow, -1, null);
            }
            if (sheet is not IStyleBinder binder)
            {
                throw new NotSupportedException($"{sheet.GetType().Name} does not support styled record layouts.");
            }
            int headerStyleId = -1;
            if (headerStyle is { } header)
            {
                headerStyleId = binder.RegisterStyle(header);
            }
            return new(headers, writeRow, headerStyleId, Register(binder, columnStyles));
        }

        /// <summary>Pushes the bound column styles, if any; the caller pops the returned binder when the rows are written.</summary>
        internal static IStyleBinder? PushOverlay<TRow>(ISheetWriter<TRow> sheet, int[]? columnStyleIds, int firstColumn)
            where TRow : IRowWriter
        {
            if (columnStyleIds is null)
            {
                return null;
            }
            var binder = (IStyleBinder)sheet;
            binder.PushColumnStyleOverlay(columnStyleIds, firstColumn);
            return binder;
        }

        private static int[]? Register(IStyleBinder binder, CellStyle?[]? styles)
        {
            if (styles is null)
            {
                return null;
            }
            int[] ids = new int[styles.Length];
            for (int i = 0; i < styles.Length; i++)
            {
                ids[i] = -1;
                if (styles[i] is { } style)
                {
                    ids[i] = binder.RegisterStyle(style);
                }
            }
            return ids;
        }
    }
}
