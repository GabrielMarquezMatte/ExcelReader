using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Mapping
{
    internal sealed unsafe class NativeSpecs : IDisposable
    {
        private NativeSpecs(XlColumnSpec* pointer, int count)
        {
            Pointer = pointer;
            Count = count;
        }

        internal XlColumnSpec* Pointer { get; private set; }

        internal int Count { get; }

        internal static NativeSpecs Create<T>(TypeMap<T> map, Dictionary<string, int> header, out PropertyBinder<T>[] bound)
            where T : new()
        {
            List<PropertyBinder<T>> matched = new List<PropertyBinder<T>>();
            List<byte[]> names = new List<byte[]>();
            foreach (PropertyBinder<T> property in map.Properties)
            {
                string? name = Array.Find(property.Names, n => header.ContainsKey(n.Trim()));
                if (name is not null)
                {
                    matched.Add(property);
                    names.Add(Utf8.Bytes(name));
                    continue;
                }

                if (property.Required)
                {
                    throw new ExcelReaderException(ExcelStatus.InvalidArgument, $"Required property '{property.PropertyName}' has no column in the header row; accepted names: {string.Join(", ", property.Names)}.");
                }
            }

            bound = matched.ToArray();
            return Allocate(bound, names);
        }

        public void Dispose()
        {
            if (Pointer is null)
            {
                return;
            }

            Marshal.FreeHGlobal((IntPtr)Pointer);
            Pointer = null;
        }

        private static NativeSpecs Allocate<T>(PropertyBinder<T>[] bound, List<byte[]> names)
        {
            int count = bound.Length;
            if (count == 0)
            {
                return new NativeSpecs(null, 0);
            }

            int textBytes = names.Sum(n => n.Length);
            int size = checked((count * (sizeof(XlColumnSpec) + sizeof(byte*) + sizeof(int))) + textBytes);
            XlColumnSpec* specs = (XlColumnSpec*)Marshal.AllocHGlobal(size);
            byte** namePointers = (byte**)(specs + count);
            int* nameLengths = (int*)(namePointers + count);
            byte* text = (byte*)(nameLengths + count);
            for (int i = 0; i < count; i++)
            {
                byte[] name = names[i];
                Marshal.Copy(name, 0, (IntPtr)text, name.Length);
                namePointers[i] = text;
                nameLengths[i] = name.Length;
                text += name.Length;

                // Always nullable: a non-nullable spec makes the native read fail on an empty or unconvertible cell, where the property should keep its default.
                specs[i] = new XlColumnSpec { Names = namePointers + i, NameLens = nameLengths + i, NameCount = 1, Index = -1, Type = bound[i].NativeType, Nullable = 1 };
            }

            return new NativeSpecs(specs, count);
        }
    }
}
