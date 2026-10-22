using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// A write-only stream that appends to a <see cref="BiffBuffer"/> it does not own. While
    /// <see cref="Discard"/> is set it drops what it is given, which is how the final block a
    /// <c>DeflateStream</c> writes on dispose is kept out of a stream of segments.
    /// </summary>
    internal sealed class BufferSink : WriteOnlyStream
    {
        private readonly BiffBuffer _target;

        internal BufferSink(BiffBuffer target)
        {
            _target = target;
        }

        internal bool Discard { get; set; }

        public override void Flush()
        {
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            Write(buffer.AsSpan(offset, count));
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            if (!Discard)
            {
                _target.Write(buffer);
            }
        }
    }
}
