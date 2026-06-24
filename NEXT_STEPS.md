# Next Steps for Another Agent

## Context

Branch: `claude/excel-parser-implementation-jvenzv`

The XLSX writer has been implemented (WorkbookWriter, SheetWriter, RowWriter) and builds
with 0 warnings. However, 13 of 20 writer round-trip tests in
`tests/ExcelReader.Tests/WorkbookWriterTests.cs` fail.

## Root Cause

The tests use **declaration-form** `await using`:

```csharp
await using RowWriter header = await sheet.StartRowAsync();
await header.WriteAsync("Name");

await using RowWriter row = await sheet.StartRowAsync();   // FAILS HERE
await row.WriteAsync("Alice");

await sheet.EndAsync();
```

In C#, `await using var x = ...;` (declaration form) disposes `x` at the END of the
enclosing block — NOT when the next statement runs. So when
`await sheet.StartRowAsync()` is called for `row`, `header` is still active
(`_rowActive == true`), causing `StartRowAsync` to throw `InvalidOperationException`.

That exception bypasses `sheet.EndAsync()`, so the sheet ZIP entry stays open.
Then `WorkbookWriter.DisposeAsync` calls `EndAsync`, which calls `CreateEntry` on an
already-open entry → `IOException: Entries cannot be created while previously
created entries are still open`.

## Fix Required

In `tests/ExcelReader.Tests/WorkbookWriterTests.cs`, convert every pair of sequential
`await using RowWriter` declarations from **declaration form** to **block form**:

```csharp
// BEFORE (broken):
await using RowWriter header = await sheet.StartRowAsync();
await header.WriteAsync("Name");

await using RowWriter row = await sheet.StartRowAsync();
await row.WriteAsync("Alice");

// AFTER (correct):
await using (RowWriter header = await sheet.StartRowAsync())
{
    await header.WriteAsync("Name");
}

await using (RowWriter row = await sheet.StartRowAsync())
{
    await row.WriteAsync("Alice");
}
```

Apply this pattern to ALL tests that write more than one row per sheet.

## Affected Tests (13 failing)

- `StringCellRoundTrip`
- `AllPrimitiveTypesRoundTrip`
- `BoolFalseRoundTrip`
- `NullStringWritesEmptyCell`
- `NullableIntFilledRoundTrip`
- `NullableDateTimeRoundTrip`
- `MultipleRowsRoundTrip`
- `MultipleSheetsRoundTrip`
- `SkipCreatesColumnGap`
- `XmlSpecialCharsAreEscaped`
- `DisposeAsyncWithoutEndAsyncProducesReadableWorkbook`
- `HeaderRowTwoConfig`
- `LargeWorkbookRoundTrip`

## Verification

```bash
dotnet build src/ExcelReader.Core/ExcelReader.Core.csproj   # must be: 0 warnings, 0 errors
dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj  # all 93 tests green
```

After all tests pass, commit and push to `claude/excel-parser-implementation-jvenzv`.
