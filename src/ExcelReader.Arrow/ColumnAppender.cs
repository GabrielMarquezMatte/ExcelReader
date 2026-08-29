using Apache.Arrow;
using Apache.Arrow.Types;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using System.Globalization;

namespace ExcelReader.Arrow
{
    /// <summary>Builds one Arrow column from a schema-driven pass over an <see cref="IExcelRowReader"/>'s rows.</summary>
    internal abstract class ColumnAppender(ExcelColumnSchema schema)
    {
        protected readonly bool Nullable = schema.IsNullable;
        protected readonly string DisplayName = schema.Name ?? schema.Index.ToString(CultureInfo.InvariantCulture);

        internal abstract Field Field { get; }

        internal abstract void Append(in Cell cell, bool isDate1904);

        internal abstract IArrowArray Build();

        /// <summary>Throws when a value failed to convert on a non-nullable column; a no-op otherwise.</summary>
        protected void ThrowIfNotNullable()
        {
            if (!Nullable)
            {
                throw new InvalidOperationException($"column \"{DisplayName}\" has a value that failed to convert and is not nullable.");
            }
        }

        internal static ColumnAppender Create(ExcelColumnSchema schema)
        {
            return schema.Type switch
            {
                ExcelColumnType.StringColumn => new StringColumnAppender(schema),
                ExcelColumnType.Int64Column => new Int64ColumnAppender(schema),
                ExcelColumnType.Float64Column => new Float64ColumnAppender(schema),
                ExcelColumnType.BoolColumn => new BoolColumnAppender(schema),
                _ => throw new NotSupportedException($"column type {schema.Type} is not supported yet."),
            };
        }

        private sealed class StringColumnAppender : ColumnAppender
        {
            private readonly StringArray.Builder _builder = new();

            internal StringColumnAppender(ExcelColumnSchema schema) : base(schema)
            {
            }

            internal override Field Field
            {
                get
                {
                    return new(DisplayName, StringType.Default, Nullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                _builder.Append(cell.GetString());
            }

            internal override IArrowArray Build()
            {
                return _builder.Build();
            }
        }

        private sealed class Int64ColumnAppender : ColumnAppender
        {
            private readonly Int64Array.Builder _builder = new();

            internal Int64ColumnAppender(ExcelColumnSchema schema) : base(schema)
            {
            }

            internal override Field Field
            {
                get
                {
                    return new(DisplayName, Int64Type.Default, Nullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (cell.TryParse(CultureInfo.InvariantCulture, out long value))
                {
                    _builder.Append(value);
                }
                else
                {
                    ThrowIfNotNullable();
                    _builder.AppendNull();
                }
            }

            internal override IArrowArray Build()
            {
                return _builder.Build();
            }
        }

        private sealed class Float64ColumnAppender : ColumnAppender
        {
            private readonly DoubleArray.Builder _builder = new();

            internal Float64ColumnAppender(ExcelColumnSchema schema) : base(schema)
            {
            }

            internal override Field Field
            {
                get
                {
                    return new(DisplayName, DoubleType.Default, Nullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (cell.TryParse(CultureInfo.InvariantCulture, out double value))
                {
                    _builder.Append(value);
                }
                else
                {
                    ThrowIfNotNullable();
                    _builder.AppendNull();
                }
            }

            internal override IArrowArray Build()
            {
                return _builder.Build();
            }
        }

        private sealed class BoolColumnAppender : ColumnAppender
        {
            private readonly BooleanArray.Builder _builder = new();

            internal BoolColumnAppender(ExcelColumnSchema schema) : base(schema)
            {
            }

            internal override Field Field
            {
                get
                {
                    return new(DisplayName, BooleanType.Default, Nullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (bool.TryParse(cell.GetString(), out bool value))
                {
                    _builder.Append(value);
                }
                else
                {
                    ThrowIfNotNullable();
                    _builder.AppendNull();
                }
            }

            internal override IArrowArray Build()
            {
                return _builder.Build();
            }
        }
    }
}
