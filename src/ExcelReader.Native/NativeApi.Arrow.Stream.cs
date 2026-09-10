using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        /// <summary>
        /// Everything one <c>ArrowArrayStream</c> needs behind its <c>private_data</c>: the batch
        /// source, the specs its schema is built from, and the last error, kept alive as UTF-8 for
        /// <c>get_last_error</c> (whose returned pointer must outlive the call).
        /// </summary>
        internal sealed class ArrowStreamSession : IDisposable
        {
            internal required TypedParseSession Session { get; init; }
            internal required NativeColumnSpec[] Specs { get; init; }

            // Named for its representation, not just its role: the outer partial class already has a
            // LastError (the xl_last_error span copy), and this is a different thing - an owned,
            // NUL-terminated UTF-8 block whose address get_last_error hands straight to the consumer.
            internal IntPtr LastErrorUtf8 { get; set; }

            internal void SetError(string message)
            {
                FreeIfSet(LastErrorUtf8);
                LastErrorUtf8 = AllocUtf8Z(message);
            }

            /// <summary>
            /// Stores the message behind a failed batch, taken from the SESSION's own latched fault
            /// rather than the thread's <c>xl_last_error</c>: that one belongs to whatever ExcelReader
            /// call ran most recently on this thread, which need not be this stream at all, so reading
            /// it back here could hand the consumer an unrelated message. Never leaves
            /// <c>get_last_error</c> empty for a non-zero return — a fault with no message of its own
            /// (nothing produces one today, but the fallback costs nothing) still gets a fixed one.
            /// </summary>
            internal void SetBatchError(string? message)
            {
                if (string.IsNullOrEmpty(message))
                {
                    message = "the Arrow stream's underlying read has faulted.";
                }
                SetError(message);
            }

            /// <summary>Closes the underlying read and frees the message block. Idempotent, because
            /// <see cref="TypedParseSession.Dispose"/> is and the freed block is zeroed here.</summary>
            public void Dispose()
            {
                // The message block is freed even if closing the read throws: TypedParseSession.Dispose
                // reaches IExcelRowEnumerator.Dispose, which can plausibly raise IOException, and the
                // caller only ever gets one contract-correct release to leak it on.
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

        /// <summary>Opens a batched Arrow read as an Arrow C Data Interface stream.</summary>
        /// <remarks>
        /// The session lives only in the stream's <c>private_data</c> as a strong (not pinned — only
        /// the handle's opaque id crosses the boundary, never the object's address)
        /// <see cref="GCHandle"/> — deliberately NOT in <see cref="NativeHandleTable"/>. There is no
        /// id for a caller to hold, so there is no stale id to misuse and no way to close this
        /// session through <c>xl_typed_reader_close</c>. <c>stream-&gt;release</c> is the only exit.
        /// </remarks>
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
                    // Allocated straight into the field: GCHandle is non-copyable (RS0042), so
                    // reading it back out of a local to pass it here would not compile.
                    PrivateData = GCHandle.ToIntPtr(GCHandle.Alloc(state)),
                };
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                // The session is open (its row enumerator holds the sheet) but nothing has been handed
                // back yet, so no `release` exists to close it and nobody else can reach it. The window
                // is narrow — a GCHandle.Alloc OOM is the only realistic trigger — but the leak would be
                // a live file handle, so it is closed here rather than left to a finalizer that
                // TypedParseSession does not have.
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

        // Arrow's get_schema/get_next return 0 on success and an errno-style code otherwise.
        private const int ArrowErrno = 5; // EIO

        internal static int ArrowStreamGetSchemaCore(ArrowArrayStream* stream, ArrowSchema* outSchema)
        {
            ArrowStreamSession? state = StateOf(stream);
            if (state is null || outSchema is null)
            {
                return ArrowErrno;
            }
            // Zeroed up front for the same reason xl_parse_arrow_stream zeroes *out_stream: a consumer
            // that ignores the errno and defensively calls out_schema->release must find a released
            // struct, not whatever its own stack happened to hold.
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
            // End of stream is a RELEASED array plus a 0 return, never an error code. Zeroing here
            // covers both that case and every failure path below.
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
                // Arrow owns an independent copy now (or, on failure, owns nothing) - the
                // intermediate batch is never reachable by the caller either way.
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
                return; // already released - Arrow permits the defensive double-release check
            }
            // The struct is marked released BEFORE anything that can throw. TypedParseSession.Dispose
            // reaches IExcelRowEnumerator.Dispose, which can plausibly raise IOException, and the thunk
            // swallows whatever escapes - so a consumer that called release exactly once, which is all
            // the contract asks of it, would otherwise be left with a leaked GCHandle and a stream
            // whose release is still non-null, i.e. one that still looks live. Reading private_data out
            // first and zeroing both fields here makes the released state unconditional; the finally
            // then guarantees the handle itself is freed on the same single call.
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
