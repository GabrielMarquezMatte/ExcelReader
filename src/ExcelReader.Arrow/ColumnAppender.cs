using Apache.Arrow;
using Apache.Arrow.Types;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using System.Globalization;

namespace ExcelReader.Arrow
{
    /// <summary>Builds one Arrow column from a schema-driven pass over an <see cref="IExcelRowReader"/>'s rows.</summary>
    internal abstract class ColumnAppender(ExcelColumnSchema schema)
    {
        protected readonly bool IsNullable = schema.IsNullable;
        protected readonly string DisplayName = schema.Name ?? schema.Index.ToString(CultureInfo.InvariantCulture);

        internal abstract Field Field { get; }

        internal abstract void Append(in Cell cell, bool isDate1904);

        internal abstract IArrowArray Build();

        /// <summary>Throws when a value failed to convert on a non-nullable column; a no-op otherwise.</summary>
        protected void ThrowIfNotNullable()
        {
            if (!IsNullable)
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
            // Cell.GetString's own stack buffer for the format-a-number branch. Matched here so an
            // unformattable number lands on the same empty result it would have through GetString.
            private const int NumberFormatMaxBytes = 32;

            private readonly StringArray.Builder _builder = new();

            // Reused across every row. Grows to the widest cell seen, never shrinks.
            private byte[] _scratch = [];

            internal StringColumnAppender(ExcelColumnSchema schema) : base(schema)
            {
            }

            internal override Field Field
            {
                get
                {
                    return new(DisplayName, StringType.Default, IsNullable);
                }
            }

            // Reading a string column always succeeds, including for a blank cell (-> ""), so it is
            // never null regardless of IsNullable.
            //
            // Cell.TryFormat emits exactly the bytes GetString would have decoded, so this hands Arrow
            // the UTF-8 it wants directly instead of allocating a managed string per cell and paying a
            // UTF-16 round trip to re-encode it. It also copies the file's bytes through unchanged
            // rather than sanitizing malformed UTF-8 to U+FFFD, matching ExcelReader.Native's
            // ColumnBuilder and every other read path in this library.
            internal override void Append(in Cell cell, bool isDate1904)
            {
                int capacity = Math.Max(cell.Value.Length, NumberFormatMaxBytes);
                if (_scratch.Length < capacity)
                {
                    _scratch = new byte[capacity];
                }
                if (!cell.TryFormat(_scratch, out int written))
                {
                    written = 0;
                }
                _builder.Append(_scratch.AsSpan(0, written));
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
                    return new(DisplayName, Int64Type.Default, IsNullable);
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
                    return new(DisplayName, DoubleType.Default, IsNullable);
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
                    return new(DisplayName, BooleanType.Default, IsNullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (ExcelCellReaders.Bool(in cell, isDate1904, CultureInfo.InvariantCulture, out bool value))
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
                    return new(DisplayName, Date32Type.Default, IsNullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (ExcelCellReaders.DateOnlyAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out DateOnly value))
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
                    return new(DisplayName, MicrosecondType, IsNullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (ExcelCellReaders.TimeOnlyAuto(in cell, isDate1904, CultureInfo.InvariantCulture, out TimeOnly value))
                {
                    _builder.Append(value.ToTimeSpan().Ticks / 10);
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
                    return new(DisplayName, MicrosecondType, IsNullable);
                }
            }

            internal override void Append(in Cell cell, bool isDate1904)
            {
                if (TryGetDateTime(in cell, isDate1904, out DateTime value))
                {
                    // Normalize whatever Kind the parse produced (Unspecified from the binary-serial path,
                    // Local from DateTime.TryParse resolving an explicit offset/Z) into a Kind-less value that
                    // represents the same absolute instant, so DateTimeOffset construction below never throws
                    // regardless of the machine's local UTC offset.
                    DateTime normalized = value.Kind == DateTimeKind.Unspecified
                        ? value
                        : DateTime.SpecifyKind(value.ToUniversalTime(), DateTimeKind.Unspecified);
                    _builder.Append(new DateTimeOffset(normalized, TimeSpan.Zero));
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
