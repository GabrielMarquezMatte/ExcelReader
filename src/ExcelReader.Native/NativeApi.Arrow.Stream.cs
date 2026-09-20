using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        internal sealed class ArrowStreamSession : IDisposable
        {
            internal required TypedParseSession Session { get; init; }
            internal required NativeColumnSpec[] Specs { get; init; }

            internal IntPtr LastErrorUtf8 { get; set; }

            internal void SetError(string message)
            {
                FreeIfSet(LastErrorUtf8);
                LastErrorUtf8 = AllocUtf8Z(message);
            }

            internal void SetBatchError(string? message)
            {
                if (string.IsNullOrEmpty(message))
                {
                    message = "the Arrow stream's underlying read has faulted.";
                }
                SetError(message);
            }

            public void Dispose()
            {
                try
                {
                    Session.Dispose();
                }
                finally
                {
                    FreeIfSet(LastErrorUtf8);
                    LastErrorUtf8 = IntPtr.Zero;
                }
            }
        }

        internal static int OpenArrowStream(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow,
            long maxRows, out ArrowArrayStream stream)
        {
            stream = default;
            int status = TypedParseSession.Open(handle, specs, headerRow, maxRows, out TypedParseSession? session);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            try
            {
                ArrowStreamSession state = new() { Session = session!, Specs = specs };
                stream = new ArrowArrayStream
                {
                    GetSchema = (IntPtr)(delegate* unmanaged<ArrowArrayStream*, ArrowSchema*, int>)&Exports.ArrowStreamGetSchema,
                    GetNext = (IntPtr)(delegate* unmanaged<ArrowArrayStream*, ArrowArray*, int>)&Exports.ArrowStreamGetNext,
                    GetLastError = (IntPtr)(delegate* unmanaged<ArrowArrayStream*, IntPtr>)&Exports.ArrowStreamGetLastError,
                    Release = (IntPtr)(delegate* unmanaged<ArrowArrayStream*, void>)&Exports.ArrowStreamRelease,
                    PrivateData = GCHandle.ToIntPtr(GCHandle.Alloc(state)),
                };
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                session!.Dispose();
                SetLastError(exception.Message);
                stream = default;
                return NativeStatus.Error;
            }
        }

        private static ArrowStreamSession? StateOf(ArrowArrayStream* stream)
        {
            if (stream is null || stream->PrivateData == IntPtr.Zero)
            {
                return null;
            }
            return GCHandle.FromIntPtr(stream->PrivateData).Target as ArrowStreamSession;
        }

        private const int ArrowErrno = 5;

        internal static int ArrowStreamGetSchemaCore(ArrowArrayStream* stream, ArrowSchema* outSchema)
        {
            ArrowStreamSession? state = StateOf(stream);
            if (state is null || outSchema is null)
            {
                return ArrowErrno;
            }
            *outSchema = default;

            try
            {
                *outSchema = BuildArrowSchema(state.Specs);
                return 0;
            }
            catch (Exception exception)
            {
                state.SetError(exception.Message);
                *outSchema = default;
                return ArrowErrno;
            }
        }

        internal static int ArrowStreamGetNextCore(ArrowArrayStream* stream, ArrowArray* outArray)
        {
            ArrowStreamSession? state = StateOf(stream);
            if (state is null || outArray is null)
            {
                return ArrowErrno;
            }
            *outArray = default;

            int status = state.Session.NextBatch(out NativeTable table);
            if (status == NativeStatus.Eof)
            {
                return 0;
            }
            if (status != NativeStatus.Ok)
            {
                state.SetBatchError(state.Session.FaultMessage);
                return ArrowErrno;
            }

            try
            {
                *outArray = BuildArrowArray(state.Specs, table);
                return 0;
            }
            catch (Exception exception)
            {
                state.SetError(exception.Message);
                *outArray = default;
                return ArrowErrno;
            }
            finally
            {
                FreeTable(ref table);
            }
        }

        internal static IntPtr ArrowStreamGetLastErrorCore(ArrowArrayStream* stream)
        {
            return StateOf(stream)?.LastErrorUtf8 ?? IntPtr.Zero;
        }

        internal static void ArrowStreamReleaseCore(ArrowArrayStream* stream)
        {
            if (stream is null || stream->Release == IntPtr.Zero)
            {
                return;
            }
            IntPtr privateData = stream->PrivateData;
            stream->PrivateData = IntPtr.Zero;
            stream->Release = IntPtr.Zero;
            if (privateData != IntPtr.Zero)
            {
                GCHandle pin = GCHandle.FromIntPtr(privateData);
                try
                {
                    (pin.Target as ArrowStreamSession)?.Dispose();
                }
                finally
                {
                    pin.Free();
                }
            }
        }
    }
}
