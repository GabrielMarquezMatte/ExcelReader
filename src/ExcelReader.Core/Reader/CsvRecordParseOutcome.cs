namespace ExcelReader.Core.Reader
{
    // Outcome of Enumerator.TryScanUnquotedRun: whether the delimiter/terminator search inside the
    // general per-field path landed on a field boundary, a record boundary, or needs more buffered bytes.
    internal enum FieldScanOutcome
    {
        NeedMore,
        FieldEnd,
        RecordEnd,
    }

    // Outcome of Enumerator.TryParseSimpleRecord: whether the fused vectorized fast path finished the
    // record, needs more buffered bytes, or bailed out because a quote turned up (general path's job).
    internal enum SimpleRecordOutcome
    {
        Done,
        NeedMore,
        Quoted,
    }
}
