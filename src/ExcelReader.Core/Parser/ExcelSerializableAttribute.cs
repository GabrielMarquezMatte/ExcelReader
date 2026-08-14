namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Marks a type for the source generator to produce <see cref="IExcelRowMap{T}"/> and
    /// <see cref="Writer.IExcelRecordMap{T}"/> implementations for, from its
    /// <c>[ExcelColumn]</c>/<c>[ExcelRequired]</c>/<c>[ExcelIgnore]</c> attributes — the same source
    /// <c>TypeMapper&lt;T&gt;</c>'s reflection path reads, but resolved at compile time. The marked type
    /// (and every type it is nested inside, if any) must be declared <c>partial</c>, since the generator
    /// emits into an additional part of the same declaration.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, Inherited = false)]
    public sealed class ExcelSerializableAttribute : Attribute
    {
    }
}
