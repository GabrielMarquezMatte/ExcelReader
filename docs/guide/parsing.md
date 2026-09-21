# Typed parsing

Binding worksheet or CSV rows to your own types: attributes, the compile-time
generator, the fluent API, converters, and the zero-copy `ref struct` path.

## Parse typed rows

`ExcelParser<T>` maps worksheet columns to the public settable properties of `T`. Columns match on the property name, or on `[ExcelColumn("header")]` aliases — repeat the attribute to accept several headers. The first row is the header by default.

```csharp
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

public sealed class ChangeRow
{
    [ExcelColumn("file")]
    public string File { get; set; } = "";

    [ExcelColumn("lines_added")]
    public int LinesAdded { get; set; }
}

using var reader = Excel.FromXlsxFile("changes.xlsx");
var parser = new ExcelParser<ChangeRow>();

foreach (var item in parser.Parse(reader))
{
    Console.WriteLine($"{item.File}: +{item.LinesAdded}");
}
```

Built-in property types: `string`, `bool`, `DateTime`, `DateOnly`, `Guid`, every integral and floating type plus `decimal`, and `enum`s (matched by member name or numeric value). Each also works as a `Nullable<T>`. Empty cells leave the property at its default; an unparseable cell is skipped (keeps the default) unless the column is required. `T` needs no parameterless-constructor constraint, so models with `required` members are supported.

`Parse` and `ParseAsync` also accept the `IExcelRowReader` from `Excel.Open`, so you can parse without knowing the concrete format:

```csharp
using IExcelRowReader reader = Excel.Open("changes.xlsx"); // or .xlsb / .xls
foreach (var item in new ExcelParser<ChangeRow>().Parse(reader)) { /* ... */ }
```

## Generate typed maps at compile time (Native AOT / trimming)

`ExcelParser<T>` and `WorkbookRecordWriter<TSheet,TRow>` reflect over `T` (`GetProperties`, `MakeGenericMethod`), which trimming can break and Native AOT cannot run at all. The raw `Excel.From*` readers already use no reflection; mark a model `[ExcelSerializable]` to get the same guarantee for the typed layer — a source generator emits a compile-time map from the model's own `[ExcelColumn]`/`[ExcelRequired]`/`[ExcelConverter]`/`[ExcelIgnore]` attributes, and `ExcelMappedParser<T>`/`MappedRecordWriter` read and write through that map instead of reflection:

```csharp
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

[ExcelSerializable]                       // the model must be declared partial
public partial class ChangeRow
{
    [ExcelColumn("file")]
    public string File { get; set; } = "";

    [ExcelColumn("lines_added")]
    public int LinesAdded { get; set; }
}

using var reader = Excel.FromXlsxFile("changes.xlsx");
foreach (var item in new ExcelMappedParser<ChangeRow>().Parse(reader))
{
    Console.WriteLine($"{item.File}: +{item.LinesAdded}");
}

var changes = new[] { new ChangeRow { File = "README.md", LinesAdded = 12 } };
await using var stream = File.Create("changes.xlsx");
await using var writer = await MappedRecordWriter.CreateMappedXlsxAsync(stream);   // or CreateMappedXlsbAsync / CreateMappedXlsAsync / CreateMappedCsvAsync
await writer.WriteSheetAsync("Changes", changes);
```

Notes:

- `[ExcelSerializable]` requires the model — and every type it's nested inside, if any — to be `partial`; the generator emits into an additional part of the same declaration. A compile error (`EXR001`/`EXR002`) names exactly what to fix.
- Supported property types match `ExcelParser<T>`'s: `string`, `bool`, `DateTime`, `DateOnly`, `TimeOnly`, `Guid`, every integral and floating type plus `decimal`, `enum`s, and `Nullable<T>` of each. Not supported: a `ref struct` model or a `ReadOnlySpan<byte>` property — those stay exclusive to `RefParser`'s reflection-based path (net9.0+); `ExcelMappedParser<T>`/`ExcelFluentParser<T>` have no AOT-clean entry for them.
- The generator requires a build via `dotnet build`/the .NET SDK. Visual Studio's or `MSBuild.exe`'s .NET Framework host can't load it, so a project built only through those tools won't see generated code — build via the SDK, or fall back to `ExcelParser<T>`/`WorkbookRecordWriter` for that build path.
- `ExcelMappedParser<T>` builds one map per model and reuses it for every reader, including CSV — unlike `ExcelParser<T>`, which swaps in a text-based date reader specifically for CSV. A `[ExcelSerializable]` model reads `DateTime`/`DateOnly`/`TimeOnly` via `ExcelCellReaders.DateTimeAuto`/`DateOnlyAuto`/`TimeOnlyAuto`: an Excel serial number first, falling back to date/time text when the cell isn't numeric — so a CSV column round-trips through the library's own writer either way. The one edge case this can't distinguish: a CSV cell that's only digits (e.g. an Excel serial number typed as plain text) is always read as a serial number, never as date text.
- The attribute-based reflection path keeps working unchanged; `[ExcelSerializable]` is an additive, opt-in alternative for the same model shape, not a replacement.

## Map columns at runtime (fluent API)

`[ExcelColumn]`/`[ExcelRequired]` fix the mapping at compile time. When the mapping itself is a runtime decision — loaded from a config file, chosen by a user in a UI, or different per input file — build it with `ExcelRowMapBuilder<T>` through `ExcelFluentParser<T>` instead:

```csharp
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

public sealed class ChangeRow
{
    public string File { get; set; } = "";
    public int LinesAdded { get; set; }
}

var parser = new ExcelFluentParser<ChangeRow>(builder => builder
    .Factory(() => new ChangeRow())
    .Property(["file"], ExcelCellReaders.String, (ref ChangeRow r, string v) => r.File = v)
    .Property(["lines_added"], ExcelCellReaders.Parsable, (ref ChangeRow r, int v) => r.LinesAdded = v));

using var reader = Excel.FromXlsxFile("changes.xlsx");
foreach (var item in parser.Parse(reader))
{
    Console.WriteLine($"{item.File}: +{item.LinesAdded}");
}
```

The map is built once, in the constructor, from a fresh builder instance — never from a static per-type cache, so two `ExcelFluentParser<T>` instances configured differently for the same `T` give different, correct results in the same process.

Bind a fixed column index instead of a header name with `PropertyAt` — for files with no header row at all:

```csharp
var parser = new ExcelFluentParser<ChangeRow>(builder => builder
    .Factory(() => new ChangeRow())
    .PropertyAt(0, ExcelCellReaders.String, (ref ChangeRow r, string v) => r.File = v)
    .PropertyAt(1, ExcelCellReaders.Parsable, (ref ChangeRow r, int v) => r.LinesAdded = v));
```

A builder that uses `PropertyAt` skips the header-row step entirely — the first row is already data. Mixing `PropertyAt` with `Property`/`PropertyNullable`/`Converted` on the same builder throws, since one map can't both wait for a header row and skip it.

`ExcelFluentParser<T>.WithAttributeFallback` merges the builder's bindings with `[ExcelColumn]`/`[ExcelRequired]`-driven ones reflected from `T`: a builder binding replaces every attribute-driven property that shares one of its header names; a property whose header names none of the builder's bindings mention keeps its attribute-driven behavior. Useful for overriding just the one column that's a runtime decision without redeclaring the whole model — reuse one of the property's existing `[ExcelColumn]` names in the builder:

```csharp
var parser = ExcelFluentParser<ChangeRow>.WithAttributeFallback(builder => builder
    .Property(["file"], ExcelCellReaders.String, (ref ChangeRow r, string v) => r.File = v.ToUpperInvariant()));
```

The match is by header name, not by property identity: configuring a *different* header name for `File` would not override its attribute — both bindings would survive and `File` would be assigned twice on the same row.

The plain constructor is AOT-clean — `configure` is caller-written code wiring hand-picked readers and setters, no reflection. `WithAttributeFallback` also reflects over `T` for the fallback half, so it carries the same `[RequiresUnreferencedCode]`/`[RequiresDynamicCode]` annotations as `ExcelParser<T>`.

## Parser configuration

Pass an `ExcelParserConfig` to control header handling and culture:

```csharp
using System.Globalization;
using ExcelReader.Core.Parser;

var config = new ExcelParserConfig
{
    HeaderRow = 1,                                   // 1-based row holding the headers
    ColumnNameComparer = StringComparer.OrdinalIgnoreCase,
    HeaderNormalization = HeaderNormalization.Trim | HeaderNormalization.CollapseSpaces,
    Culture = CultureInfo.GetCultureInfo("pt-BR"),   // parse "1.234,56" as 1234.56m
};

var parser = new ExcelParser<ChangeRow>(config);
```

`Culture` applies when parsing text-backed numeric/`Guid` cells (XLSX inline and shared strings); binary numeric cells (XLS/XLSB) carry a raw value and ignore it. `HeaderNormalization` flags (`Trim`, `CollapseSpaces`, `RemoveDiacritics`) are applied to both the sheet headers and the property names before matching.

## Required columns

Mark a property `[ExcelRequired]` to assert its column exists and carries a value:

```csharp
public sealed class Order
{
    [ExcelRequired]
    public int Id { get; set; }

    [ExcelRequired(AllowEmpty = true)]   // column must exist; blank cells allowed
    public string? Note { get; set; }
}
```

- A missing required header throws when the header row is read, listing every missing column.
- By default each data row must have a non-empty cell; the first blank throws, naming the column and row number. `AllowEmpty = true` relaxes this to column presence only.
- The check covers presence, not parseability — a present-but-malformed value does not throw here.

## Custom converters

For types the built-in parsers do not handle — money strings, custom formats, domain value objects — implement `IExcelCellConverter<T>` and attach it with `[ExcelConverter]`. `T` must be the property's exact type. One instance is created and reused across all rows, so converters must be stateless.

```csharp
using System.Globalization;
using ExcelReader.Core.Parser;
using ExcelReader.Core.ValueObjects;

public sealed class BrlMoneyConverter : IExcelCellConverter<decimal>
{
    public bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out decimal value)
    {
        string text = cell.GetString().Replace("R$", "", StringComparison.Ordinal).Trim();
        return decimal.TryParse(text, NumberStyles.Currency, CultureInfo.GetCultureInfo("pt-BR"), out value);
    }
}

public sealed class Invoice
{
    [ExcelConverter(typeof(BrlMoneyConverter))]
    public decimal Total { get; set; }
}
```

Return `false` to signal a parse failure (the property keeps its default). Empty cells are skipped before the converter runs.

## Parse into a ref struct (zero-copy)

`RefParser.ParseNamed<T>` (.NET 9+) targets a `ref struct` model instead of a class/struct — same attribute-driven column matching as `ExcelParser<T>` (`[ExcelColumn]`, `[ExcelRequired]`, `[ExcelConverter]`), but a `ReadOnlySpan<byte>` property binds directly to the cell's raw bytes instead of allocating a `string`:

```csharp
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

public readonly ref struct ChangeRowRef
{
    public ReadOnlySpan<byte> File { get; init; }   // zero-copy — aliases the reader's row buffer
    public int LinesAdded { get; init; }
}

using var reader = Excel.FromXlsxFile("changes.xlsx");

foreach (ChangeRowRef item in RefParser.ParseNamed<ChangeRowRef>(reader))
{
    Console.WriteLine($"{Encoding.UTF8.GetString(item.File)}: +{item.LinesAdded}");
}
```

The sequence also supports `await foreach`, so a `ref struct` model can be parsed asynchronously — the rows are streamed via `MoveNextAsync` while the model stays a zero-copy `ref struct`:

```csharp
await using var reader = await Excel.FromXlsxFileAsync("changes.xlsx");

await foreach (ChangeRowRef item in RefParser.ParseNamed<ChangeRowRef>(reader))
{
    Console.WriteLine($"{Encoding.UTF8.GetString(item.File)}: +{item.LinesAdded}");
}
```

A few differences from `ExcelParser<T>`:

- **Span fields alias the reader's row buffer** — valid only until the next row. Copy them out (e.g. `Encoding.UTF8.GetString(span)`) if you need to keep the value past the loop body. Under `await foreach`, the same rule means the model can't be held across an `await` in the loop body.
- **`foreach` / `await foreach` only.** Consumption is pattern-based — the sequence cannot be surfaced through `IEnumerable<T>`, `IAsyncEnumerable<T>`, or LINQ, because a `ref struct` element can't be boxed through those interfaces (`IAsyncEnumerable<T>` in particular forbids a `ref struct` element type — CS9267). Iterate it directly.
- **Not AOT/trim-safe**, same tradeoff as `ExcelParser<T>` (both reflect over `T`'s properties and compile setters at runtime).
- A regular `struct`/`class` model works with `ParseNamed` too — only a genuine `ref struct` model gets the extra zero-copy span-property binding.

