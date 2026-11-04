namespace ExcelReader.NetFramework.Reader
{
    /// <summary>
    /// A fold over the rows of a sheet, driven by <see cref="ExcelSheet.Aggregate{TAccumulator}"/>. One instance is
    /// seeded per chunk of sheet data, not per thread, and a chunk can be accumulated and then dropped without ever
    /// being combined. <see cref="Accumulate"/> runs concurrently on worker threads, each call on its own instance.
    /// </summary>
    /// <typeparam name="TSelf">The implementing type.</typeparam>
    public interface IExcelAccumulator<TSelf>
        where TSelf : IExcelAccumulator<TSelf>
    {
        /// <summary>Folds one row in.</summary>
        /// <param name="row">The row, valid only for the duration of the call.</param>
        void Accumulate(ExcelRowView row);

        /// <summary>Folds the contents of a later chunk into this one; calls run in source order.</summary>
        /// <param name="other">The later chunk, discarded once the call returns.</param>
        void Combine(TSelf other);
    }
}
