# ExcelReader.NET.Framework

Fast XLSX, XLSB, XLS and CSV reading for .NET Framework 4.7.2 and later. A thin P/Invoke binding over
the ExcelReader NativeAOT library, with no managed package dependencies and no Open XML SDK.

## Install

    Install-Package ExcelReader.NET.Framework

or `dotnet add package ExcelReader.NET.Framework`. The package copies `ExcelReader.Native.dll` into an
`x64` folder of the build output. On first use the library looks for `x64\ExcelReader.Native.dll`
under each of these folders, in order:

1. the application base directory (`AppDomain.BaseDirectory`);
2. each private bin path (`AppDomain.RelativeSearchPath`, e.g. `bin` for an ASP.NET site);
3. the original folder of `ExcelReader.NetFramework.dll` (`Assembly.CodeBase`, before shadow copying);
4. the folder `ExcelReader.NetFramework.dll` was loaded from (`Assembly.Location`).

## x64 process required in this version

Only the x64 native library ships in this version. A 32-bit process fails on first use with a
`PlatformNotSupportedException` that lists every path it probed. Make sure the process is 64-bit:

```xml
<PropertyGroup>
  <PlatformTarget>x64</PlatformTarget>
</PropertyGroup>
```

`AnyCPU` works too, as long as **Prefer 32-bit** is off (`<Prefer32Bit>false</Prefer32Bit>`) and the
process runs on a 64-bit OS. In Visual Studio: Project Properties, Build, Platform target.

## Read a sheet into objects

```csharp
using ExcelReader.NetFramework.Parser;
using ExcelReader.NetFramework.Reader;

public sealed class Invoice
{
    [ExcelColumn("Invoice No")]
    [ExcelRequired]
    public string Number { get; set; }

    public DateTime? Issued { get; set; }

    [ExcelColumn("Total")]
    [ExcelColumn("Amount")]
    public decimal Total { get; set; }

    [ExcelIgnore]
    public bool Checked { get; set; }
}

using (ExcelWorkbook workbook = ExcelWorkbook.Open(@"C:\data\invoices.xlsx"))
{
    Invoice[] invoices = workbook.Sheet("Invoices").Read<Invoice>();
}
```

`Read<T>(headerRow = 1, degreeOfParallelism = 0)` reads every data row. For big sheets,
`ReadBatches<T>(batchRows)` yields arrays of at most `batchRows` items so the whole sheet never sits in
memory. `ReadRows()` enumerates raw cells, `InferSchema()` reports column types, `Table(name)` reads an
Excel table through the same calls, and `ExcelWorkbook.Open` also takes a `byte[]` or a `Stream`.

Namespaces mirror ExcelReader.Core's folders: `ExcelReader.NetFramework.Reader` holds the workbook types, `.Reader.Schema` the schema types and `.Parser` the mapping attributes.

Header matching trims the header text, ignores case, and the first matching column wins. A property
whose column is absent, or a cell that is empty or does not convert, keeps its default value unless the
property is `[ExcelRequired]`.

## Attributes

| Attribute | Effect |
|---|---|
| `[ExcelColumn("Header")]` | Maps the property to the column with that header. Repeat it to accept alternate spellings. Without it the property name is the header. |
| `[ExcelRequired]` | The column must be in the header row, and every data cell must be non-empty and parseable. Parsing throws `ExcelReaderException` otherwise. |
| `[ExcelRequired(AllowEmpty = true)]` | Only the column's presence is required. |
| `[ExcelIgnore]` | The property is not read. |

Only public properties with a setter are read, on a class with a public parameterless constructor. A
property without a setter is skipped, unless it has `[ExcelColumn]` or `[ExcelRequired]`: then the read
throws `InvalidOperationException`.

## Supported property types

| Property type | Notes |
|---|---|
| `string` | Cell text. |
| `long`, `int`, `short`, `byte` | Whole numbers; a value outside the type's range throws. |
| `double`, `float`, `decimal` | Numbers. `decimal` is converted from the double the native reader produces, so digits beyond about 15 significant figures are lost. |
| `bool` | Boolean cells. |
| `DateTime` | Date and date-time cells. |
| `TimeSpan` | Time-of-day cells. |

Each value type also works as `Nullable<T>` (`int?`, `DateTime?`, ...), which stays `null` for an empty
cell. Any other type, including enums and structs, throws `NotSupportedException`.

## Walk or fold rows without copying

```csharp
using ExcelReader.NetFramework.Reader;

public sealed class Total : IExcelAccumulator<Total>
{
    public double Sum;

    public void Accumulate(ExcelRowView row)
    {
        if (row.Count > 1 && row[1].Number is double value)
        {
            Sum += value;
        }
    }

    public void Combine(Total other)
    {
        Sum += other.Sum;
    }
}

using (ExcelWorkbook workbook = ExcelWorkbook.Open(@"C:\data\sales.xlsx"))
{
    ExcelSheet sheet = workbook.Sheet(0);
    double sum = sheet.Aggregate(() => new Total()).Sum;
    sheet.ForEach(row => Console.WriteLine(row[0].GetText()));
}
```

`ForEach(action)` calls `action` once per row, header included, on the calling thread. `Aggregate(seed,
headerRow = 1, degreeOfParallelism = 0)` skips the header and every row above it, seeds one accumulator
per chunk of the sheet on worker threads, and combines them in sheet order. XLS, tables and a CSV opened
from a `Stream` are read sequentially, with the same result. An exception thrown by a callback ends the
call and is rethrown to the caller.

An `ExcelRowView` and its `ExcelCellView`s are valid only during the callback that received them: they
cannot be stored, and `ToRow()` copies one into an `ExcelRow`. `Number` is the native double, or the
text parsed with the invariant culture. `GetText()` allocates the string, and formats the number when the
cell holds only a number.

## Write

```csharp
using ExcelReader.NetFramework.Reader;
using ExcelReader.NetFramework.Writer;

// Typed: a list becomes a sheet. The format comes from the extension.
using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.Create(@"\\server\share\sales.xlsx"))
{
    writer.WriteSheet("Sales", sales, new ExcelTableOptions { Name = "Sales" });
    writer.Close();
}

// Download: build in memory, copy to the response
using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx))
{
    writer.WriteSheet("Sales", sales);
    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    Response.AddHeader("Content-Disposition", "attachment; filename=vendas.xlsx");
    writer.CloseTo(Response.OutputStream);   // or byte[] file = writer.CloseToArray();
}

// Cell by cell, mixed with a typed sheet
using (ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx))
{
    writer.StartSheet("Summary");
    writer.StartRow();
    writer.Write("Total");
    writer.Write(1234.5);
    writer.EndRow();
    writer.EndSheet();

    writer.WriteSheet("Sales", sales);

    byte[] file = writer.CloseToArray();
}
```

`Create(path, options)` picks the format from `.xlsx`, `.xlsb`, `.xls` or `.csv` (or
`ExcelWriteOptions.Format`); `CreateInMemory(format, options)` writes to memory. `Close()` finishes a path
writer; `CloseToArray()` and `CloseTo(Stream)` finish an in-memory one and do not close the stream.
`Abort()` discards everything. **`Dispose` without `Close` also discards the output**: no file appears at
a new path and an existing file is left intact, so an exception mid-generation never leaves a half-written
workbook.

Cell methods: `Write(string)`, `Write(long)`, `Write(double)`, `Write(decimal)`, `Write(bool)`,
`WriteDate`, `WriteTime`, `WriteDateTime` and `WriteEmpty`. `int`, `short` and `byte` go through
`Write(long)`. `WriteSheet<T>` is only valid between sheets, not while one is open.

`WriteSheet<T>` maps the same property types as reading (see above); `decimal` is written as `double`, with the same
precision loss as reading. The header is
the first `[ExcelColumn]` name or the property name, `[ExcelIgnore]` properties are skipped, and columns
follow declaration order, base class first. A `null` writes a blank cell. `DateTime` writes date and
time, so a date-only column shows `00:00`. `TimeSpan` must be within `[0, 24h)`.

- CSV cannot tell an empty string from `null`: both write an empty field. XLSX, XLSB and XLS keep `""`
  as an empty string cell, distinct from `null` (no cell). CSV has no sheet names; read one back with
  `Sheet(0)`.
- `ExcelTableOptions` (Excel tables) need XLSX or XLSB; other formats throw `ExcelReaderException`.
  `ExcelWriteOptions.Password` also needs XLSX or XLSB and throws `NotSupportedException` otherwise; an
  empty password throws `ArgumentException`. A path writer with a password builds in memory and writes the encrypted
  bytes on `Close`.
- No cell formatting, formulas or charts.

## Threading

Calls on one `ExcelWorkbook` can run on several threads at once, and `Read<T>` itself parses on
several threads (`degreeOfParallelism`, default: every processor). A single `ReadRows()` or
`ReadBatches<T>()` enumeration belongs to one thread at a time. A CSV opened from a `Stream` is read as
it arrives, so it allows one row cursor at a time. Dispose the workbook when you are done; enumerations
already running keep the native data they need alive until they finish.

## More

Source, benchmarks and the other bindings (.NET, Python, C): https://github.com/GabrielMarquezMatte/ExcelReader
