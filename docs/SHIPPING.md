# Shipping Checklist

Scope: read-only raw data extraction from XLSX. No write support, formula eval, or rich formatting — intentionally out of scope.

---

## Bugs

### 1. 1904 date system — silent data corruption
Mac-authored workbooks store `<workbookPr date1904="1"/>` in `workbook.xml`. Without detecting it every date is off by 1,462 days — no error, just wrong data.

Fix: read the flag in `ParseSheets`, pass it through to `TryGetDateTime`, add 1462 to the serial if set.

### 2. No `TryGetBool()` on `Cell`
`CellType.Boolean` is detected correctly from `t="b"` but the API forces the consumer to call `GetString()` and manually parse `"0"` / `"1"`. Add:
```csharp
public bool TryGetBool(out bool value)
```

### 3. Column index overflow on malformed files
`XlsxXml.ColumnIndex` silently wraps on column references longer than ~8 letters. Add a guard after each multiply:
```csharp
if (col > 16_384) { return -1; }
```

### 4. `SheetName` expression body — styleguide violation
`public string SheetName => _sheets[_current].Name;` must be a block body per STYLEGUIDE §3.

---

## Missing features (within scope)

### 5. No sheet enumeration API
`SheetCount` and index-based `MoveToSheet` are fine, but there is no way to get all sheet names. Add:
```csharp
public IReadOnlyList<string> SheetNames { get; }
```

### 6. Error cell value undocumented
`CellType.Error` is detected and the raw text (`#DIV/0!`, `#N/A`, etc.) is in `Value`, but it is not documented. `GetString()` works — at minimum add a `<remarks>` comment. Optionally add `TryGetErrorText(out string text)` for symmetry.

### 7. Enumerator is forward-only — not documented
Once a sheet is fully iterated, rewinding requires calling `GetEnumerator()` again (opens a new `ZipArchiveEntry` stream). This surprises consumers used to `IEnumerable`. Document it clearly on `GetEnumerator()` and `GetAsyncEnumeratorAsync()`.

---

## Packaging (blocking)

### 8. NuGet metadata missing
Add to `ExcelReader.Core.csproj`:
```xml
<Version>1.0.0</Version>
<Authors>...</Authors>
<Description>Fast, zero-dependency, streaming XLSX reader with async support.</Description>
<PackageLicenseExpression>MIT</PackageLicenseExpression>
<PackageTags>excel;xlsx;reader;streaming;async;performance</PackageTags>
<PackageReadmeFile>README.md</PackageReadmeFile>
<PackageProjectUrl>...</PackageProjectUrl>
<RepositoryUrl>...</RepositoryUrl>
<RepositoryType>git</RepositoryType>
```

### 9. No LICENSE file
Add `LICENSE` (MIT recommended for a utility library) to the repo root.

### 10. No README.md
Minimum content: what it does, a 10-line sync + async usage example, known limitations (read-only, XLSX only, 1900 date system until bug #1 is fixed).

### 11. Target framework decision
`GetAsyncEnumeratorAsync` uses `ZipArchive.CreateAsync` which is .NET 10 only. Options:
- Keep `net10.0` only and document the requirement clearly.
- Multi-target `net8.0;net9.0;net10.0` by falling back to sync zip open inside `Task.Run` on older frameworks.

### 12. No XML doc comments
`Cell`, `Row`, `CellType`, `Excel`, `XlsxReader` have no `<summary>` comments. IDE shows nothing on hover. One-line summaries on all public members is the minimum bar.

---

## Priority

| # | Item | Effort | Blocks ship? |
|---|------|--------|-------------|
| 8–10 | Metadata + LICENSE + README | 1–2 h | Yes |
| 2 | `TryGetBool()` | 30 min | Yes |
| 5 | `SheetNames` property | 20 min | Yes |
| 12 | XML doc comments | 2 h | Yes |
| 11 | .NET targeting decision | varies | Yes |
| 1 | 1904 date fix | 1 h | Strongly recommended |
| 3 | Column overflow guard | 10 min | No |
| 4 | `SheetName` block body | 5 min | No |
| 6–7 | Error cell + rewind docs | 30 min | No |
