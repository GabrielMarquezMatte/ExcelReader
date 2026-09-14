namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        /// <summary>
        /// Opens a batch-at-a-time typed read over the current sheet and registers it, returning the
        /// opaque id <c>xl_typed_reader_next</c>/<c>xl_typed_reader_close</c> take.
        /// </summary>
        /// <remarks>
        /// The id comes from <see cref="NativeHandleTable"/>, so it is never recycled and is
        /// type-checked on resolve: a stale reader id stays permanently invalid rather than naming
        /// some later workbook or writer.
        /// </remarks>
        internal static int OpenTypedReader(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow,
            long maxRows, out nint reader)
        {
            reader = 0;
            int status = TypedParseSession.Open(handle, specs, headerRow, maxRows, out TypedParseSession? session);
            if (status != NativeStatus.Ok)
            {
                return status;
            }
            reader = NativeHandleTable.Register(session!);
            return NativeStatus.Ok;
        }

        internal static int NextTypedBatch(nint reader, out NativeTable table)
        {
            table = default;
            TypedParseSession? session = NativeHandleTable.Resolve<TypedParseSession>(reader);
            if (session is null)
            {
                SetLastError("the typed reader id is not a live reader.");
                return NativeStatus.InvalidHandle;
            }
            return session.NextBatch(out table);
        }

        internal static void CloseTypedReader(nint reader)
        {
            if (NativeHandleTable.TryUnregister(reader, out TypedParseSession? session))
            {
                session!.Dispose();
            }
        }
    }
}
