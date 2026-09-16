namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
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
