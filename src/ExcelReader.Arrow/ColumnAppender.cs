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

        /// <summary>
        /// Reads a date/time value from either an Excel serial (binary formats) or plain text (CSV,
        /// which has no serial encoding); <see cref="DateTime"/> doesn't implement
        /// <see cref="IUtf8SpanParsable{TSelf}"/>, so the text fallback goes through <see cref="Cell.GetString"/>.
        /// </summary>
        protected static bool TryGetDateTime(in Cell cell, bool isDate1904, out DateTime value)
        {
            return cell.TryGetDateTime(isDate1904, out value)
                || DateTime.TryParse(cell.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
        }

        internal static ColumnAppender Create(ExcelColumnSchema schema)
        {
            return schema.Type switch
            {
                ExcelColumnType.StringColumn => new StringColumnAppender(schema),
                ExcelColumnType.Int64Column => new Int64ColumnAppender(schema),
                ExcelColumnType.Float64Column => new Float64ColumnAppender(schema),
                ExcelColumnType.BoolColumn => new BoolColumnAppender(schema),
                ExcelColumnType.DateColumn => new DateColumnAppender(schema),
                ExcelColumnType.TimeColumn => new TimeColumnAppender(schema),
                ExcelColumnType.TimestampColumn => new TimestampColumnAppender(schema),
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

        private sealed class DateColumnAppender : ColumnAppender
        {
            private readonly Date32Array.Builder _builder = new();

            internal DateColumnAppender(ExcelColumnSchema schema) : base(schema)
            {
            }

            internal override Field Field
            {
                get
                {
                    return new(DisplayName, Date32Type.Default, Nullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (TryGetDateTime(in cell, isDate1904, out DateTime value))
                {
                    _builder.Append(value.Date);
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

        private sealed class TimeColumnAppender : ColumnAppender
        {
            private static readonly Time64Type MicrosecondType = new(TimeUnit.Microsecond);

            private readonly Time64Array.Builder _builder = new(MicrosecondType);

            internal TimeColumnAppender(ExcelColumnSchema schema) : base(schema)
            {
            }

            internal override Field Field
            {
                get
                {
                    return new(DisplayName, MicrosecondType, Nullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (TryGetDateTime(in cell, isDate1904, out DateTime value))
                {
                    _builder.Append(value.TimeOfDay.Ticks / 10);
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

        private sealed class TimestampColumnAppender : ColumnAppender
        {
            // No timezone: matches the native `xl_parse_arrow` path's "tsu:" format code.
            private static readonly TimestampType MicrosecondType = new(TimeUnit.Microsecond, timezone: (string?)null);

            private readonly TimestampArray.Builder _builder = new(MicrosecondType);

            internal TimestampColumnAppender(ExcelColumnSchema schema) : base(schema)
            {
            }

            internal override Field Field
            {
                get
                {
                    return new(DisplayName, MicrosecondType, Nullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (TryGetDateTime(in cell, isDate1904, out DateTime value))
                {
                    // Zero offset keeps the raw naive value as microseconds since epoch.
                    _builder.Append(new DateTimeOffset(value, TimeSpan.Zero));
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
