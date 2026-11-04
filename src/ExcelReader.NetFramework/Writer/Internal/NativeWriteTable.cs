using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Writer.Internal
{
    internal sealed unsafe class NativeWriteTable : IDisposable
    {
        private readonly List<IntPtr> _blocks = new List<IntPtr>();

        private NativeWriteTable()
        {
        }

        internal XlTable* Table { get; private set; }

        internal XlColumnSpec* Specs { get; private set; }

        internal static NativeWriteTable Create<T>(WriteMap<T> map, T[] rows)
        {
            NativeWriteTable table = new NativeWriteTable();
            try
            {
                table.Build(map, rows);
                return table;
            }
            catch
            {
                table.Dispose();
                throw;
            }
        }

        internal byte* Allocate(long bytes)
        {
            _blocks.Add(IntPtr.Zero);
            IntPtr block = Marshal.AllocHGlobal(new IntPtr(Math.Max(bytes, 1)));
            _blocks[_blocks.Count - 1] = block;
            return (byte*)block;
        }

        internal byte* AllocateValidity(int rows)
        {
            int length = (rows + 7) / 8;
            byte* validity = Allocate(length);
            for (int i = 0; i < length; i++)
            {
                validity[i] = 0;
            }

            return validity;
        }

        public void Dispose()
        {
            foreach (IntPtr block in _blocks)
            {
                if (block != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(block);
                }
            }

            _blocks.Clear();
            Table = null;
            Specs = null;
        }

        private void Build<T>(WriteMap<T> map, T[] rows)
        {
            ColumnWriter<T>[] writers = map.Columns;
            int count = writers.Length;
            XlColumnSpec* specs = (XlColumnSpec*)Allocate((long)count * sizeof(XlColumnSpec));
            XlColumn* columns = (XlColumn*)Allocate((long)count * sizeof(XlColumn));
            for (int i = 0; i < count; i++)
            {
                specs[i] = Spec(writers[i], i);
                writers[i].Fill(rows, columns + i, this);
            }

            XlTable* table = (XlTable*)Allocate(sizeof(XlTable));
            *table = new XlTable { ColumnCount = count, RowCount = rows.Length, Columns = columns };
            Specs = specs;
            Table = table;
        }

        private XlColumnSpec Spec<T>(ColumnWriter<T> writer, int index)
        {
            byte[] name = Utf8.Bytes(writer.Header);
            byte** names = (byte**)Allocate(sizeof(byte*));
            int* nameLengths = (int*)Allocate(sizeof(int));
            byte* text = Allocate(name.Length);
            Marshal.Copy(name, 0, (IntPtr)text, name.Length);
            *names = text;
            *nameLengths = name.Length;
            return new XlColumnSpec { Names = names, NameLens = nameLengths, NameCount = 1, Index = index, Type = writer.NativeType, Nullable = writer.Nullable ? 1 : 0 };
        }
    }
}
