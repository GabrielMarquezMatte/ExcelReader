# Versioning and compatibility

What a new version can change, and how to write code that keeps compiling and running across minor
versions. The .NET packages follow [Semantic Versioning](https://semver.org/): a major version (7.0,
8.0) may break your code, while a minor (7.1) or a patch (7.0.1) does not.

## What a minor version may add

- New types, new members on existing types, and new overloads.
- New properties on the options objects, such as
  [`ExcelReaderOptions`](xref:ExcelReader.Core.Reader.ExcelReaderOptions). A new option always defaults
  to the previous behavior.
- New members on the interfaces that only this library implements (below).
- New values in the enums (below).
- Bug fixes. A file that used to read wrong, or a written file that Excel refused or altered, can come
  out differently after a fix. Other changes in what an existing call returns or throws wait for a
  major version and are listed in its migration guide.

A minor version does not remove or rename anything, change an existing signature, or add a
parameter to an existing method.

## Interfaces you implement

These are the extension points. A minor version never adds a member to them that you would have to
implement:

- [`IExcelCellConverter<T>`](xref:ExcelReader.Core.Parser.IExcelCellConverter`1) and
  [`IExcelCellWriter<T>`](xref:ExcelReader.Core.Writer.IExcelCellWriter`1)
- [`IExcelRowMap<T>`](xref:ExcelReader.Core.Parser.IExcelRowMap`1) and
  [`IExcelRecordMap<T>`](xref:ExcelReader.Core.Writer.IExcelRecordMap`1)
- [`IRowRecord<TSelf>`](xref:ExcelReader.Core.Reader.IRowRecord`1) and
  [`IRowAccumulator<TSelf, TModel>`](xref:ExcelReader.Core.Reader.IRowAccumulator`2)

New information reaches them through the types they already receive. For example, a converter gets
it as a new member of [`ExcelRowContext`](xref:ExcelReader.Core.Parser.ExcelRowContext), not as a
new parameter.

## Interfaces only the library implements

The reading and writing interfaces describe what every format's workbook, sheet and row writer can
do. Program against them, but do not implement them:

- [`IExcelWorkbook`](xref:ExcelReader.Core.Reader.IExcelWorkbook),
  [`IExcelSheet`](xref:ExcelReader.Core.Reader.IExcelSheet) and
  [`IExcelSheet<TEnumerator>`](xref:ExcelReader.Core.Reader.IExcelSheet`1)
- [`IExcelRowEnumerator`](xref:ExcelReader.Core.Reader.IExcelRowEnumerator)
- [`IWorkbookWriter<TSheet>`](xref:ExcelReader.Core.Writer.IWorkbookWriter`1),
  [`ISheetWriter<TRow>`](xref:ExcelReader.Core.Writer.ISheetWriter`1),
  [`ITableSheetWriter<TRow>`](xref:ExcelReader.Core.Writer.ITableSheetWriter`1) and
  [`IRowWriter`](xref:ExcelReader.Core.Writer.IRowWriter)

A minor version can add members to them. Each new member comes with a default implementation, so an
existing implementation still compiles and loads, but it throws `NotSupportedException` or falls back
to a slower path for the new member. A feature that only some formats support appears first on the
concrete types (`XlsxSheetWriter`, `XlsbWorkbook`), or as a separate interface the way
[`ITableSheetWriter<TRow>`](xref:ExcelReader.Core.Writer.ITableSheetWriter`1) sits beside
[`ISheetWriter<TRow>`](xref:ExcelReader.Core.Writer.ISheetWriter`1).

## Enums that grow

[`CellType`](xref:ExcelReader.Core.Reader.CellType) and
[`ExcelColumnType`](xref:ExcelReader.Core.Reader.Schema.ExcelColumnType) can gain values in a minor
version, as can their native counterparts `XL_CELL_*` and `XL_T_*`. An existing value never changes
its number or meaning. A `switch` over them should have a default branch:

<!-- snippet: versioning-enum-switch -->
```csharp
string text = cell.Type switch
{
    CellType.Empty => "",
    CellType.Number when cell.TryGetDouble(out double number) => number.ToString(CultureInfo.InvariantCulture),
    _ => cell.GetString(),
};
```

## The native library and the bindings

The C ABI is versioned on its own by `XL_ABI_VERSION`, which `xl_abi_version()` reports. It changes
only when a signature or a struct layout does: 7.0 kept version 6. Every options struct starts with
`struct_size`, so a field added in a later version does not break a caller compiled against an
earlier header. The Rust, Python and C++ bindings check the ABI version when they load the library.
`xl_write_options` grew from 32 to 40 bytes after 7.0 (`degree_of_parallelism` and `batch_rows`,
appended); the 7.0 size is still accepted and writes serially, so the ABI version stayed 6.

7.x also added `xl_rows_foreach` and `xl_aggregate`. They are new exports only, so the ABI version
stayed 6; a binding built against them fails with a missing-symbol error when it loads an older
ABI-6 library that lacks them.

## How it is enforced

Every pull request packs `ExcelReader.NET` and `ExcelReader.Arrow` and compares them with the last
release through .NET package validation. A removed or changed member fails the build.
