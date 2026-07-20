namespace ExcelReader.Core.Parser.Internal
{
    // Outcome of feeding one source row to the projector: skip it (header/pre-header rows),
    // yield the projected model, or stop iterating (bindings never built).
    internal enum ProjectionStep
    {
        Skip,
        BuildMap,
        Yield,
        Stop,
    }
}
