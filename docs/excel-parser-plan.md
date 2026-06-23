# Excel Parser Implementation Plan

## Context

The project has a low-level, high-performance XLSX reader (`XlsxReader`) that streams rows as `Row` (ref struct) and cells as `Cell` (ref struct). This plan adds a parser layer that maps rows to POCO types (`T`) using compiled reflection — modeled after the FixedWidthParser library pattern.

Key design decisions:
- **No `yield return`** — custom struct enumerators
- **TryParse pattern** — `ColumnParser<T>` delegate returns `bool`, no exceptions in hot path
- **Expression-tree-compiled setters** (`RefAction<T, TProp>`) — works for both classes and structs
- **Regular lambdas for column parsers** — close over the compiled setter

---

## File Structure

```
src/ExcelReader.Core/
  Parser/
    ExcelColumnAttribute.cs       # [ExcelColumn("Name")]
    ExcelParserConfig.cs          # HeaderRow, ColumnNameComparer
    ExcelParser.cs                # public API: Parse / ParseAsync / TryParse
    Internal/
      Delegates.cs                # RefAction<TModel,TProperty>, ColumnParser<TModel>
      ColumnParserFactory.cs      # builds ColumnParser<T> per property type
      TypeMapper.cs               # property discovery + static lazy cache
      ExcelEnumerable.cs          # IEnumerable<T> + struct Enumerator
      ExcelAsyncEnumerable.cs     # IAsyncEnumerable<T> + async enumerator class

tests/ExcelReader.Tests/
  ExcelParserTests.cs
```

---

## Public API

### `ExcelColumnAttribute`

Maps a property to a named Excel column header. When absent, the property name is used.

### `ExcelParserConfig`

- `HeaderRow` (int, default 1): 1-based row number of the column headers.
- `ColumnNameComparer` (StringComparer, default `OrdinalIgnoreCase`): controls how header names match property names.

### `ExcelParser<T>` (where T : new())

- `Parse(XlsxReader)` → `ExcelEnumerable<T>` — sync, struct enumerator, no allocation per call.
- `ParseAsync(XlsxReader, CancellationToken)` → `ExcelAsyncEnumerable<T>` — async, no yield.
- `TryParse(ref T, in Row, ColumnParser<T>?[], bool)` → bool — low-level, shared by both paths.

---

## Internal Architecture

### Delegate Types

```csharp
// Setter compiled once per property via Expression tree.
// ref TModel allows in-place mutation for both classes and structs.
internal delegate void RefAction<TModel, in TProperty>(ref TModel model, TProperty value)
    where TModel : allows ref struct;

// Column-level TryParse: accesses row[columnIndex] internally.
// Returns false on parse failure; true on success or empty cell.
internal delegate bool ColumnParser<TModel>(
    ref TModel model, in Row row, int columnIndex, bool isDate1904)
    where TModel : allows ref struct;
```

### `ColumnParserFactory`

1. Build `RefAction<T, TProp>` via `Expression.Lambda` with `typeof(T).MakeByRefType()` parameter — works for both classes and structs.
2. Build `ColumnParser<T>` as a regular C# lambda closing over the setter.

Type dispatch:

| Property type | Conversion |
|---------------|------------|
| `string` | `cell.GetString()` |
| `int`, `long`, `double`, `float`, `decimal` | `cell.TryParse<T>(null, out T)` — UTF-8 direct |
| `DateTime` | `cell.TryGetDateTime(isDate1904, out dt)` |
| `bool` | `cell.Value.SequenceEqual("1"u8)` |
| `Nullable<TInner>` | Empty → return true (null stays); else inner conversion |

### `TypeMapper<T>`

Static `Lazy<TypeMapInfo<T>>` on a generic static class — one instance per T, zero dictionary overhead.

### Enumerators

- **Sync**: `ExcelEnumerable<T>` returns a `struct Enumerator`. `MoveNext()` is a plain while-loop with no yield.
- **Async**: `ExcelAsyncEnumerable<T>` uses a sealed `AsyncEnumerator` class. The `async ValueTask<bool> AdvanceAsync()` method only awaits; row access (ref struct) is delegated to a synchronous `ProcessCurrentRow()` helper so ref structs never cross an await boundary.

---

## Pitfalls

1. `ref T` Expression tree: use `typeof(T).MakeByRefType()` for the model parameter.
2. Ref struct in async: only access `_rows.Current` inside sync helpers, never in the async method.
3. Nullable: empty cell → `return true`, inner parse failure → `return false`.
4. Array bounds: `Math.Min(_maxColumn, row.ColumnCount)` for data rows.
5. Build errors captured via `ExceptionDispatchInfo` in `TypeMapper`, rethrown at call site.
6. Bool parsing: compare bytes directly, avoid `GetString()`.
