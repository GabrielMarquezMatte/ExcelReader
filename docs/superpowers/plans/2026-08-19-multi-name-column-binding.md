# Multi-name column binding (native bindings) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a native-binding column spec (Rust `#[excel(...)]`, C++ `make_field`, Python `ColumnSpec`) carry an ordered list of candidate header names, resolved first-match-wins against the header row — the native-binding equivalent of the C# core's `[ExcelColumn(AllowMultiple = true)]`.

**Architecture:** The single point of truth is `src/ExcelReader.Native` (a C# NativeAOT project compiled to the shared library every binding links against). `xl_column_spec` moves from one `name`/`name_len` pair to an ordered `names`/`name_lens`/`name_count` array; the header-resolution loop in `NativeApi.Typed.cs` tries each candidate in order and keeps the first one present in the header row. This is a breaking ABI change (`XL_ABI_VERSION` 1 → 2), so every binding that mirrors `xl_column_spec` — C header, C++ header, Rust FFI struct, Python ctypes struct — is updated in lockstep, and every existing single-name call site keeps compiling/working unchanged (a plain string/one name is just a one-element list).

**Tech Stack:** C# (.NET, NativeAOT), C17 header, C++23, Rust (stable, `syn`/`quote` proc-macro), Python 3 (`ctypes`).

**Spec:** `docs/superpowers/specs/2026-08-19-multi-name-column-binding-design.md`

## Global Constraints

- `XL_ABI_VERSION` bumps from `1` to `2` in all four places it is mirrored: `src/ExcelReader.Native/include/excelreader.h`, `src/ExcelReader.Native/NativeStatus.cs`, `rust/excelreader/src/lib.rs`, `python/src/excelreader/_native.py`. Miss one and `python/tests/test_native.py::test_abi_version_matches_header` (or the Rust/C++ equivalents) fails.
- Header matching stays case-insensitive and trimmed — unchanged rule, now applied per candidate name instead of once.
- A **write** spec (`xl_write_typed`) and an **inferred** column (`xl_infer_schema`) always carry exactly 0 or 1 names — only the **read** path (`xl_parse_typed`/`xl_parse_arrow`) resolves a candidate list of more than one.
- Every existing single-name call site (C# `Name = "x"`, C++ `make_field("x", member)`, Rust `#[excel(name = "x")]`, Python `ColumnSpec(name="x")`) must keep compiling and behaving identically — the multi-name form is additive.

---

## Task 1: C ABI struct + C# native core read path

**Files:**
- Modify: `src/ExcelReader.Native/include/excelreader.h:24` (XL_ABI_VERSION), `excelreader.h:200-208` (xl_column_spec)
- Modify: `src/ExcelReader.Native/NativeStatus.cs:16` (AbiVersion)
- Modify: `src/ExcelReader.Native/NativeTypedTable.cs:26-46` (NativeColumnSpecRaw, NativeColumnSpec)
- Modify: `src/ExcelReader.Native/Exports.cs:401-421` (TryDecodeColumnSpecs)
- Modify: `src/ExcelReader.Native/NativeApi.Typed.cs:58-59,145-245` (error message, TryValidateArguments, TryResolveColumns, FindHeaderColumn)
- Test: `tests/ExcelReader.Tests/NativeApiTests.cs`

**Interfaces:**
- Produces: `NativeColumnSpec.Names` (`string[]`, empty = index-based, replaces `Name`), consumed by every later task's C# work (Task 2) and read directly by every existing/new test in this task.

- [ ] **Step 1: Write the failing tests**

Add to `tests/ExcelReader.Tests/NativeApiTests.cs`, near the other `ParseTyped` tests (e.g. after the existing `ParseTyped_Should_Reject_An_Unmatched_Header_Name` test around line 1941-1961, whose `File.WriteAllText`/`try`/`finally` shape these two mirror exactly):

```csharp
[Fact]
public void ParseTyped_Should_Resolve_The_First_Candidate_Name_Present_In_The_Header()
{
    string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
    File.WriteAllText(path, "qty,quantity\n5\n");
    try
    {
        Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
        NativeColumnSpec[] specs = [new() { Names = ["does-not-exist", "quantity"], Type = NativeColumnType.Int64 }];

        int status = NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table);

        Assert.Equal(NativeStatus.Ok, status);
        Assert.Equal(1, table.RowCount);
        long value = Marshal.ReadInt64(ColumnAt(table, 0).Values);
        Assert.Equal(5, value);
        NativeApi.FreeTable(ref table);
        NativeApi.Close(handle);
    }
    finally
    {
        File.Delete(path);
    }
}

[Fact]
public void ParseTyped_Should_Fail_With_A_Message_Listing_Every_Candidate_When_None_Match()
{
    string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
    File.WriteAllText(path, "qty\n5\n");
    try
    {
        Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
        NativeColumnSpec[] specs = [new() { Names = ["nope", "still-nope"], Type = NativeColumnType.Int64 }];

        int status = NativeApi.ParseTyped(handle, specs, headerRow: 1, out NativeTable table);

        Assert.Equal(NativeStatus.InvalidArgument, status);
        Assert.Equal(IntPtr.Zero, table.Columns);
        Span<byte> buffer = stackalloc byte[256];
        NativeApi.LastError(buffer, out int length);
        string message = Encoding.UTF8.GetString(buffer[..length]);
        Assert.Contains("\"nope\"", message);
        Assert.Contains("\"still-nope\"", message);
        NativeApi.Close(handle);
    }
    finally
    {
        File.Delete(path);
    }
}
```

`ColumnAt` (the existing private helper at `NativeApiTests.cs:1582`) and `Encoding` (`System.Text`, already used elsewhere in this file for UTF-8 assertions) are both already available in this test class — no new helpers needed.

- [ ] **Step 2: Run the new tests to verify they fail to compile**

Run: `dotnet test tests/ExcelReader.Tests --filter "FullyQualifiedName~ParseTyped_Should_Resolve_The_First_Candidate|FullyQualifiedName~ParseTyped_Should_Fail_With_A_Message_Listing_Every_Candidate"`
Expected: build error — `NativeColumnSpec` has no `Names` member yet (only `Name`).

- [ ] **Step 3: Bump XL_ABI_VERSION and change the C header struct**

In `src/ExcelReader.Native/include/excelreader.h`, change line 24:

```c
#define XL_ABI_VERSION 2
```

Replace lines 200-208:

```c
/* Describes one output column of xl_parse_typed. Resolve by header name (names != NULL, matched
 * case-insensitively and trimmed against the header row, trying each candidate in `names` in order
 * and stopping at the first match) or by physical column index (name_count == 0). */
typedef struct xl_column_spec {
    const uint8_t* const* names; /* candidate header texts, UTF-8, in priority order; NULL when
                                   * name_count == 0 to match by index instead */
    const int32_t* name_lens;    /* one length per entry in `names` */
    int32_t name_count;
    int32_t index;                /* zero-based column index, used when name_count == 0 */
    int32_t type;                 /* XL_T_* */
    int32_t nullable;             /* 0 = a failed conversion is XL_ERROR; 1 = it becomes null (validity bit 0) */
} xl_column_spec;
```

- [ ] **Step 4: Bump NativeStatus.AbiVersion**

In `src/ExcelReader.Native/NativeStatus.cs:16`:

```csharp
internal const int AbiVersion = 2;
```

- [ ] **Step 5: Change NativeTypedTable.cs's raw and decoded specs**

Replace lines 26-46 of `src/ExcelReader.Native/NativeTypedTable.cs`:

```csharp
/// <summary>
/// Flat C ABI representation for one raw <c>xl_column_spec</c> as received across the boundary — the
/// <see cref="Names"/> pointers are only valid for the duration of the call, so <see cref="Exports"/>
/// decodes them into the UTF-8-decoded <see cref="NativeColumnSpec"/> before calling into
/// <see cref="NativeApi"/>.
/// </summary>
[StructLayout(LayoutKind.Sequential)]
internal unsafe struct NativeColumnSpecRaw
{
    public byte** Names;
    public int* NameLens;
    public int NameCount;
    public int Index;
    public int Type;
    public int Nullable;
}

/// <summary>Decoded, pointer-free form of <see cref="NativeColumnSpecRaw"/> — the layer
/// <see cref="NativeApi"/> and its tests actually work with.</summary>
internal readonly struct NativeColumnSpec
{
    /// <summary>Candidate header texts to match (case-insensitively, trimmed), tried in order — the
    /// first one present in the header row wins. Empty to resolve by <see cref="Index"/> instead.</summary>
    internal string[] Names { get; init; } = [];
    internal int Index { get; init; }
    internal int Type { get; init; }
    internal bool Nullable { get; init; }
}
```

- [ ] **Step 6: Decode the array in Exports.cs**

Replace `TryDecodeColumnSpecs` (lines 401-421 of `src/ExcelReader.Native/Exports.cs`):

```csharp
// Shared by xl_parse_typed and xl_parse_arrow, whose column-spec input is identical. Returns
// false for a name length that cannot describe a real header rather than passing it to
// GetString, where it would become a read length over however much caller memory it names.
private static bool TryDecodeColumnSpecs(NativeColumnSpecRaw* specs, int specCount, out NativeColumnSpec[] decoded)
{
    decoded = new NativeColumnSpec[specCount];
    for (int i = 0; i < specCount; i++)
    {
        NativeColumnSpecRaw raw = specs[i];
        if (raw.NameCount < 0 || (raw.NameCount > 0 && (raw.Names is null || raw.NameLens is null)))
        {
            decoded = [];
            return false;
        }
        string[] names = new string[raw.NameCount];
        for (int n = 0; n < raw.NameCount; n++)
        {
            if (!NativeApi.IsValidNameLength(raw.NameLens[n]))
            {
                decoded = [];
                return false;
            }
            names[n] = Encoding.UTF8.GetString(raw.Names[n], raw.NameLens[n]);
        }
        decoded[i] = new NativeColumnSpec
        {
            Names = names,
            Index = raw.Index,
            Type = raw.Type,
            Nullable = raw.Nullable != 0,
        };
    }
    return true;
}
```

- [ ] **Step 7: Rewrite validation and resolution in NativeApi.Typed.cs**

Replace the failing-column error message on line 59:

```csharp
NativeColumnSpec spec = specs[failedColumn];
string columnLabel = spec.Names.Length > 0 ? string.Join(" / ", spec.Names) : spec.Index.ToString(CultureInfo.InvariantCulture);
SetLastError($"column {failedColumn} (\"{columnLabel}\") has a value that failed to convert and is not nullable.");
```

Replace `TryValidateArguments` (lines 145-184):

```csharp
private static bool TryValidateArguments(NativeColumnSpec[] specs, int headerRow, out string? error)
{
    error = null;
    if (specs.Length == 0)
    {
        error = "xl_parse_typed requires at least one column spec.";
        return false;
    }
    if (headerRow < 0)
    {
        error = $"header_row must be 0 (no header) or a positive row number; got {headerRow}.";
        return false;
    }
    foreach (NativeColumnSpec spec in specs)
    {
        if (spec.Names.Length == 0 && spec.Index < 0)
        {
            error = "a column spec with no name must have a non-negative index.";
            return false;
        }
        if (spec.Names.Length > 0 && headerRow == 0)
        {
            error = $"column \"{spec.Names[0]}\" is name-based, but header_row is 0 (no header row to match it against).";
            return false;
        }
        // FindHeaderColumn trims before comparing, so a blank name would match the first empty
        // header cell — resolving to a column the caller never asked for instead of failing.
        foreach (string name in spec.Names)
        {
            if (name.AsSpan().Trim().IsEmpty)
            {
                error = "a name-based column spec cannot have a blank name.";
                return false;
            }
        }
        if (spec.Type is < NativeColumnType.String or > NativeColumnType.Timestamp)
        {
            error = $"column spec has unknown type {spec.Type}.";
            return false;
        }
    }
    return true;
}
```

Replace `TryResolveColumns` (lines 189-227):

```csharp
// Advances `rows` past any skipped rows and the header row itself (headerRow > 0), or leaves it
// untouched at the sheet's first row (headerRow == 0, index-only specs). Either way, `rows` is
// positioned so the next MoveNext() yields the first DATA row.
private static bool TryResolveColumns(IExcelRowEnumerator rows, NativeColumnSpec[] specs, int headerRow, int[] columnIndices, out string? error)
{
    error = null;
    if (headerRow == 0)
    {
        for (int i = 0; i < specs.Length; i++)
        {
            columnIndices[i] = specs[i].Index;
        }
        return true;
    }

    for (int rowNumber = 1; rowNumber <= headerRow; rowNumber++)
    {
        if (!rows.MoveNext())
        {
            error = $"sheet has fewer than {headerRow} row(s); cannot resolve header_row.";
            return false;
        }
    }

    Row header = rows.Current;
    for (int i = 0; i < specs.Length; i++)
    {
        string[] names = specs[i].Names;
        if (names.Length == 0)
        {
            columnIndices[i] = specs[i].Index;
            continue;
        }
        int found = -1;
        foreach (string name in names)
        {
            found = FindHeaderColumn(header, name);
            if (found >= 0)
            {
                break;
            }
        }
        if (found < 0)
        {
            error = $"no column header matches any of {FormatCandidates(names)}.";
            return false;
        }
        columnIndices[i] = found;
    }
    return true;
}

private static string FormatCandidates(string[] names)
{
    return string.Join(", ", Array.ConvertAll(names, n => $"\"{n}\""));
}
```

`FindHeaderColumn` (lines 234-245) is unchanged — it already takes one name and is now called once per candidate.

- [ ] **Step 8: Bulk-rename the existing single-name test call sites**

Every existing `NativeColumnSpec` object initializer in `tests/ExcelReader.Tests/NativeApiTests.cs` uses `Name = "..."` or `Name = null`. Confirmed by grep: all ~48 occurrences of `\bName = ` in that file are `NativeColumnSpec` initializers (no false positives from sheet names or other properties). Apply this mechanical rewrite:

```bash
sed -i -E 's/Name = "([^"]*)"/Names = ["\1"]/g; s/Name = null/Names = []/g' tests/ExcelReader.Tests/NativeApiTests.cs
```

This does **not** touch `DecodeSchema`'s local tuple field `spec.Name` at (originally) line 2219 — that is a test-local `(string? Name, ...)` tuple returned by the `DecodeSchema` helper, unrelated to `NativeColumnSpec`, and is handled separately in Task 2.

- [ ] **Step 9: Run the full NativeApiTests.cs suite**

Run: `dotnet test tests/ExcelReader.Tests --filter "FullyQualifiedName~NativeApiTests"`
Expected: every `ParseTyped`/`WriteTyped`/`ParseArrow` test that only reads specs still passes unchanged (single-element `Names` arrays behave exactly like the old single `Name`); the two new tests from Step 1 pass. `WriteTyped`/`InferSchema`-related tests will still fail to compile at this point — that's expected, they're fixed in Task 2.

- [ ] **Step 10: Commit**

```bash
git add src/ExcelReader.Native/include/excelreader.h src/ExcelReader.Native/NativeStatus.cs src/ExcelReader.Native/NativeTypedTable.cs src/ExcelReader.Native/Exports.cs src/ExcelReader.Native/NativeApi.Typed.cs tests/ExcelReader.Tests/NativeApiTests.cs
git commit -m "feat(native): resolve xl_parse_typed columns against a candidate name list

Bumps XL_ABI_VERSION to 2. xl_column_spec.name/name_len become
names/name_lens/name_count; the read path (xl_parse_typed/xl_parse_arrow)
tries each candidate in order and binds the first one present in the
header row."
```

---

## Task 2: C# native core write path + schema-inference output

**Files:**
- Modify: `src/ExcelReader.Native/NativeApi.Write.cs:157-166,285-319`
- Modify: `src/ExcelReader.Native/NativeApi.Schema.cs:77-95,184-202`
- Test: `tests/ExcelReader.Tests/NativeApiTests.cs`

**Interfaces:**
- Consumes: `NativeColumnSpec.Names` (`string[]`), `NativeColumnSpecRaw { Names, NameLens, NameCount, Index, Type, Nullable }` from Task 1.
- Produces: nothing new consumed by later tasks — this task closes out the C# side.

- [ ] **Step 1: Write the failing tests**

Add to `tests/ExcelReader.Tests/NativeApiTests.cs`, near the other `WriteTyped` validation tests:

```csharp
[Fact]
public void WriteTyped_Should_Reject_A_Spec_With_More_Than_One_Name()
{
    // A well-formed 1-column, 1-row table: the point of this test is that TryValidateWriteTable
    // rejects the name_count > 1 spec specifically, not that it rejects a malformed table for some
    // unrelated reason first (e.g. a column-count mismatch).
    NativeTable table = BuildInt64Table([1L]);
    try
    {
        NativeColumnSpec[] specs = [new() { Names = ["qty", "quantity"], Type = NativeColumnType.Int64 }];
        string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
        try
        {
            int status = NativeApi.WriteTyped(Encoding.UTF8.GetBytes(path), NativeFormat.Csv, specs, table, new NativeWriteOptions());
            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.False(File.Exists(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }
    finally
    {
        FreeBuiltTable(ref table);
    }
}
```

Update `DecodeSchema`'s return type and body (originally lines 2228-2244) — this is test-support code, not new coverage, but it must read the new raw layout before any `InferSchema` test compiles:

```csharp
private static (string? Name, int Index, int Type, bool Nullable)[] DecodeSchema(NativeInferredSchema schema)
{
    int specSize = Marshal.SizeOf<NativeColumnSpecRaw>();
    int namesOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.Names));
    int nameLensOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.NameLens));
    int nameCountOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.NameCount));
    int indexOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.Index));
    int typeOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.Type));
    int nullableOffset = (int)Marshal.OffsetOf<NativeColumnSpecRaw>(nameof(NativeColumnSpecRaw.Nullable));

    var columns = new (string?, int, int, bool)[schema.ColumnCount];
    for (int i = 0; i < columns.Length; i++)
    {
        IntPtr spec = IntPtr.Add(schema.Columns, i * specSize);
        int nameCount = Marshal.ReadInt32(spec, nameCountOffset);
        string? name = null;
        if (nameCount > 0)
        {
            IntPtr namesArray = Marshal.ReadIntPtr(spec, namesOffset);
            IntPtr nameLensArray = Marshal.ReadIntPtr(spec, nameLensOffset);
            IntPtr namePtr = Marshal.ReadIntPtr(namesArray, 0);
            int nameLen = Marshal.ReadInt32(nameLensArray, 0);
            name = Marshal.PtrToStringUTF8(namePtr, nameLen);
        }
        int index = Marshal.ReadInt32(spec, indexOffset);
        int type = Marshal.ReadInt32(spec, typeOffset);
        int nullable = Marshal.ReadInt32(spec, nullableOffset);
        columns[i] = (name, index, type, nullable != 0);
    }
    return columns;
}
```

`AssertSpec` (originally line 2217) is unchanged — it destructures this same local tuple, whose `Name` field name never changed.

- [ ] **Step 2: Run the new/updated tests to verify they fail to compile**

Run: `dotnet test tests/ExcelReader.Tests --filter "FullyQualifiedName~NativeApiTests"`
Expected: build errors — `WriteHeaderRow` still reads `spec.Name` (removed in Task 1), `TryValidateWriteTable`/`BuildSpec`/`FreeSchema` still reference the old raw/decoded shape.

- [ ] **Step 3: Fix the write path in NativeApi.Write.cs**

Replace `WriteHeaderRow` (lines 157-166):

```csharp
private static void WriteHeaderRow<TSheet, TRow>(TSheet sheet, NativeColumnSpec[] specs)
    where TSheet : ISheetWriter<TRow>
    where TRow : IRowWriter
{
    using TRow row = sheet.StartRow();
    foreach (NativeColumnSpec spec in specs)
    {
        row.Write(spec.Names[0]);
    }
}
```

Replace `TryValidateWriteTable` (lines 285-319):

```csharp
internal static bool TryValidateWriteTable(NativeColumnSpec[] specs, NativeTable table, out bool hasHeader, [NotNullWhen(false)] out string? error)
{
    hasHeader = false;
    error = null;
    if (specs.Length == 0 || !IsValidSpecCount(table.ColumnCount))
    {
        error = $"xl_write_typed needs 1..{NativeLimits.MaxColumnSpecs} columns; got {table.ColumnCount}.";
        return false;
    }
    if (specs.Length != table.ColumnCount)
    {
        error = $"xl_write_typed got {specs.Length} spec(s) for {table.ColumnCount} column(s).";
        return false;
    }
    if (table.RowCount < 0 || table.Columns == IntPtr.Zero)
    {
        error = $"xl_write_typed needs a non-negative row_count and a non-NULL columns pointer; got {table.RowCount}.";
        return false;
    }

    hasHeader = specs[0].Names.Length > 0;
    for (int index = 0; index < table.ColumnCount; index++)
    {
        int nameCount = specs[index].Names.Length;
        if ((nameCount > 0) != hasHeader)
        {
            error = "every column spec must have a name, or none may — xl_write_typed cannot write a partial header row.";
            return false;
        }
        if (nameCount > 1)
        {
            error = $"column {index} is a write spec and must have exactly one name; got {nameCount}.";
            return false;
        }
        if (!TryValidateWriteColumn(specs[index], ColumnAt(table, index), index, table.RowCount, out error))
        {
            return false;
        }
    }
    return true;
}
```

`TryValidateWriteColumn` is unchanged — it never reads `.Name`/`.Names`.

- [ ] **Step 4: Fix the schema-inference output in NativeApi.Schema.cs**

Replace `FreeSchema` (lines 77-95):

```csharp
internal static void FreeSchema(ref NativeInferredSchema schema)
{
    if (schema.Columns == IntPtr.Zero)
    {
        schema = default;
        return;
    }

    NativeColumnSpecRaw* columns = (NativeColumnSpecRaw*)schema.Columns;
    for (int i = 0; i < schema.ColumnCount; i++)
    {
        NativeColumnSpecRaw spec = columns[i];
        if (spec.NameCount > 0)
        {
            Marshal.FreeHGlobal((IntPtr)spec.Names[0]);
            Marshal.FreeHGlobal((IntPtr)spec.Names);
            Marshal.FreeHGlobal((IntPtr)spec.NameLens);
        }
    }
    Marshal.FreeHGlobal(schema.Columns);
    schema = default;
}
```

Replace `BuildSpec` (lines 184-202):

```csharp
private static NativeColumnSpecRaw BuildSpec(string? name, int index, ColumnStat stat)
{
    byte** namesBlock = null;
    int* lensBlock = null;
    int nameCount = 0;
    if (name is not null)
    {
        int nameLen = Encoding.UTF8.GetByteCount(name);
        byte* namePtr = (byte*)Marshal.AllocHGlobal(Math.Max(nameLen, 1));
        Encoding.UTF8.GetBytes(name, new Span<byte>(namePtr, nameLen));

        namesBlock = (byte**)Marshal.AllocHGlobal(sizeof(byte*));
        namesBlock[0] = namePtr;
        lensBlock = (int*)Marshal.AllocHGlobal(sizeof(int));
        lensBlock[0] = nameLen;
        nameCount = 1;
    }
    return new NativeColumnSpecRaw
    {
        Names = namesBlock,
        NameLens = lensBlock,
        NameCount = nameCount,
        Index = index,
        Type = stat.InferType(),
        Nullable = stat.SawEmpty ? 1 : 0,
    };
}
```

- [ ] **Step 5: Run the full test project**

Run: `dotnet test tests/ExcelReader.Tests`
Expected: PASS, including `ExcelReader.NativeSmoke`'s own build if it's part of the same solution build (see Task 3 for that project specifically).

- [ ] **Step 6: Commit**

```bash
git add src/ExcelReader.Native/NativeApi.Write.cs src/ExcelReader.Native/NativeApi.Schema.cs tests/ExcelReader.Tests/NativeApiTests.cs
git commit -m "feat(native): adapt write and schema-inference paths to the candidate-name spec shape

Both remain single-name (write rejects name_count > 1; inference always
emits 0 or 1 names), now expressed through names/name_lens/name_count."
```

---

## Task 3: C smoke test

**Files:**
- Modify: `tests/ExcelReader.NativeSmoke/smoke.c:58-63,446-772`

**Interfaces:**
- Consumes: the `xl_column_spec` shape from Task 1, Step 3.

This is a build-verification file (compiled and run against the real native library in CI), not TDD in the red/green sense — there is no separate assertion framework to drive first. The verification step is "it compiles and the CHECKs still pass."

- [ ] **Step 1: Update the struct-layout static asserts**

Replace lines 58-63:

```c
XL_STATIC_ASSERT(offsetof(xl_column_spec, names) == 0, column_spec_names);
XL_STATIC_ASSERT(offsetof(xl_column_spec, name_lens) == 8, column_spec_name_lens);
XL_STATIC_ASSERT(offsetof(xl_column_spec, name_count) == 16, column_spec_name_count);
XL_STATIC_ASSERT(offsetof(xl_column_spec, index) == 20, column_spec_index);
XL_STATIC_ASSERT(offsetof(xl_column_spec, type) == 24, column_spec_type);
XL_STATIC_ASSERT(offsetof(xl_column_spec, nullable) == 28, column_spec_nullable);
XL_STATIC_ASSERT(sizeof(xl_column_spec) == 32, column_spec_size);
```

- [ ] **Step 2: Add a one-name helper**

Add near the top of the file, after the existing `#include`s:

```c
/* Fills `spec` as a single-candidate name-based spec, using `name_slot`/`len_slot` as the
 * one-element backing storage `spec->names`/`spec->name_lens` point into — that storage must
 * outlive every use of `spec` (the caller declares it in the same or an outer scope). */
static void set_spec_name1(xl_column_spec* spec, const uint8_t** name_slot, int32_t* len_slot, const char* text)
{
    *name_slot = (const uint8_t*)text;
    *len_slot = (int32_t)strlen(text);
    spec->names = name_slot;
    spec->name_lens = len_slot;
    spec->name_count = 1;
}
```

- [ ] **Step 3: Update `build_specs` (originally lines 446-458)**

```c
static int build_specs(xl_column_spec* specs, const uint8_t** name_ptrs, int32_t* name_lens)
{
    memset(specs, 0, 3 * sizeof(xl_column_spec));
    set_spec_name1(&specs[0], &name_ptrs[0], &name_lens[0], "Coluna1");
    specs[0].type = XL_T_STRING;
    specs[0].nullable = 1;
    set_spec_name1(&specs[1], &name_ptrs[1], &name_lens[1], "Coluna2");
    specs[1].type = XL_T_DATE;
    specs[1].nullable = 1;
    set_spec_name1(&specs[2], &name_ptrs[2], &name_lens[2], "Coluna3");
    specs[2].type = XL_T_I64;
    specs[2].nullable = 1;
    return 3;
}
```

Update every call site of `build_specs(specs)` to also declare and pass the backing storage, e.g.:

```c
xl_column_spec specs[3];
const uint8_t* name_ptrs[3];
int32_t name_lens[3];
build_specs(specs, name_ptrs, name_lens);
```

- [ ] **Step 4: Update every remaining single-spec construction site**

For each of the remaining sites (originally around lines 474-477, 496-501, 516-519, 754-772), replace the two-line `X.name = ...; X.name_len = ...;` pattern with a sibling storage pair and a `set_spec_name1` call. For example, the `one_spec` fixture (originally lines 474-477):

```c
xl_column_spec one_spec;
memset(&one_spec, 0, sizeof(one_spec));
const uint8_t* one_spec_name;
int32_t one_spec_name_len;
set_spec_name1(&one_spec, &one_spec_name, &one_spec_name_len, "Coluna1");
```

and the `wide_name`/`blank_name` variants derived from it (originally lines 496-519) keep their `xl_column_spec wide_name = one_spec;` copy — a struct copy also copies the `names`/`name_lens` pointers, which still point at `one_spec_name`/`one_spec_name_len`, so overriding just `wide_name.name_lens` requires its own backing slot instead:

```c
xl_column_spec wide_name = one_spec;
int32_t wide_name_len = XL_MAX_COLUMN_NAME_BYTES + 1;
wide_name.name_lens = &wide_name_len;
/* ... */
wide_name_len = -1;
```

```c
xl_column_spec blank_name;
memset(&blank_name, 0, sizeof(blank_name));
const uint8_t* blank_name_name;
int32_t blank_name_len;
set_spec_name1(&blank_name, &blank_name_name, &blank_name_len, "   ");
```

And the `spec` used for the write-typed section (originally lines 754/771-772):

```c
xl_column_spec spec;
memset(&spec, 0, sizeof(spec));
const uint8_t* spec_name;
int32_t spec_name_len;
set_spec_name1(&spec, &spec_name, &spec_name_len, "qty");
```

- [ ] **Step 5: Update the inferred-schema read sites (originally lines 666-676)**

```c
xl_column_spec coluna1 = schema.columns[0];
CHECK(coluna1.name_count == 1 && coluna1.name_lens[0] == 7 && memcmp(coluna1.names[0], "Coluna1", 7) == 0, "column 0 must be named Coluna1");
```

(same transformation for `coluna2`/`coluna3`).

- [ ] **Step 6: Build and run the smoke test**

Run: `cmake --build build --target excelreader_native_smoke && ./build/tests/ExcelReader.NativeSmoke/excelreader_native_smoke` (adjust the exact target/binary path to whatever the existing CMake configuration names it — check `tests/ExcelReader.NativeSmoke/CMakeLists.txt` or the root build if unsure).
Expected: every `CHECK` line passes (exit code 0), against a freshly built native library carrying Task 1/2's changes (`EXCELREADER_NATIVE_LIB` pointed at a local build, per `cpp/cmake/FetchNativeLib.cmake`'s override).

- [ ] **Step 7: Commit**

```bash
git add tests/ExcelReader.NativeSmoke/smoke.c
git commit -m "test(native-smoke): update the C smoke test for the candidate-name xl_column_spec layout"
```

---

## Task 4: C++ binding

**Files:**
- Modify: `src/ExcelReader.Native/include/excelreader.hpp:338-376,466-499`
- Test: `cpp/tests/smoke.cpp`

**Interfaces:**
- Consumes: the `xl_column_spec` shape from Task 1, Step 3.
- Produces: `xl::make_field({...}, member)` (multi-name overload) and `xl::FieldBinding<Class, T, N>`, for any downstream C++ consumer code (none inside this repo beyond `cpp/tests`).

- [ ] **Step 1: Write the failing test**

Add to `cpp/tests/smoke.cpp`, a new mapper and test function exercising alias resolution against the real `RealExcel.xlsb` fixture (whose header is `Coluna1`):

```cpp
struct AliasRow {
    std::string_view Coluna1;
};

template<> struct xl::ExcelMapper<AliasRow> {
    static constexpr auto get_bindings() {
        return std::make_tuple(
            xl::make_field({"ThisColumnDoesNotExist", "Coluna1"}, &AliasRow::Coluna1)
        );
    }
};

static int test_parse_with_alias(xl::Workbook& workbook) {
    auto table = xl::parse_sheet<AliasRow>(workbook);
    CHECK(table.has_value(), "xl::parse_sheet<AliasRow> must succeed by resolving the second candidate name");
    CHECK(table->size() == 100, "RealExcel.xlsb has 100 data rows");
    Row first = *table->begin();
    return 0;
}
```

Wire it into `main()` alongside the existing `test_parse(workbook)` call. (`Row` in the last line is a typo guard — replace with reading `AliasRow`'s own first element, e.g. `AliasRow first = *table->begin(); CHECK(first.Coluna1 == "Valor1", "first row's Coluna1 must be Valor1");`.)

- [ ] **Step 2: Build to verify it fails to compile**

Run: `cmake --build build --target excelreader_smoke` (check `cpp/CMakeLists.txt`/`cpp/tests` for the exact target name if different)
Expected: compile error — `make_field` has no overload taking a braced list as its first argument yet.

- [ ] **Step 3: Update FieldBinding and make_field (lines 466-480)**

```cpp
// ---- Struct <-> column bindings ---------------------------------------------------------------

template <typename Class, typename T, std::size_t N = 1>
struct FieldBinding
{
    std::array<const char *, N> column_names;
    T Class::*member;
    using FieldType = T;
};

template <typename Class, typename T>
constexpr FieldBinding<Class, T, 1> make_field(const char *name, T Class::*member)
{
    return {{name}, member};
}

// N is deduced from the braced-init-list argument via array-reference binding, e.g.
// make_field({"Nome", "Nom", "Name"}, &Row::nome).
template <typename Class, typename T, std::size_t N>
constexpr FieldBinding<Class, T, N> make_field(const char *(&&names)[N], T Class::*member)
{
    FieldBinding<Class, T, N> result{};
    for (std::size_t i = 0; i < N; ++i)
    {
        result.column_names[i] = names[i];
    }
    result.member = member;
    return result;
}
```

- [ ] **Step 4: Update build_specs (lines 486-499)**

```cpp
namespace detail
{
    template <typename Class, typename T, std::size_t N>
    xl_column_spec build_one_spec(const FieldBinding<Class, T, N> &binding, std::vector<int32_t> &name_lens_storage)
    {
        name_lens_storage.resize(N);
        for (std::size_t i = 0; i < N; ++i)
        {
            name_lens_storage[i] = static_cast<int32_t>(std::strlen(binding.column_names[i]));
        }
        return xl_column_spec{
            reinterpret_cast<const uint8_t *const *>(binding.column_names.data()),
            name_lens_storage.data(),
            static_cast<int32_t>(N),
            0, // index is ignored: resolved by name
            XlType<T>::value,
            1 // nullable = 1 (safe default)
        };
    }

    template <typename Tuple, std::size_t... Is>
    std::array<xl_column_spec, sizeof...(Is)> build_specs(
        const Tuple &bindings,
        std::index_sequence<Is...>,
        std::array<std::vector<int32_t>, sizeof...(Is)> &name_lens_storage)
    {
        return {build_one_spec(std::get<Is>(bindings), name_lens_storage[Is])...};
    }
```

(the rest of the `detail` namespace — `is_valid`, `assign_field`, `populate_instance` — is unchanged; keep them where they are relative to this block).

- [ ] **Step 5: Update parse_sheet's call site (lines 763-780)**

```cpp
template <typename T>
std::expected<TableView<T>, Error> parse_sheet(Workbook &workbook, int32_t header_row = 1)
{
    static constexpr auto bindings = ExcelMapper<T>::get_bindings();
    static constexpr size_t num_fields = std::tuple_size_v<decltype(bindings)>;

    std::array<std::vector<int32_t>, num_fields> name_lens_storage{};
    std::array<xl_column_spec, num_fields> specs_array =
        detail::build_specs(bindings, std::make_index_sequence<num_fields>{}, name_lens_storage);
    std::span<const xl_column_spec> specs(specs_array);

    xl_table table{};
    int32_t status = xl_parse_typed(workbook.handle(), specs.data(), static_cast<int32_t>(specs.size()), header_row, &table);
    if (status != XL_OK)
    {
        return std::unexpected(detail::make_error(status));
    }
    return TableView<T>::from_raw(table);
}
```

- [ ] **Step 6: Update infer_schema's decode (lines 357-374)**

```cpp
std::vector<InferredColumn> columns;
columns.reserve(static_cast<size_t>(schema.column_count > 0 ? schema.column_count : 0));
for (int32_t i = 0; i < schema.column_count; ++i)
{
    const xl_column_spec &spec = schema.columns[i];
    InferredColumn column{};
    // A guessed name is exactly name_lens[0] bytes with no NUL terminator, and name_count is 0
    // whenever the column had no usable header cell.
    if (spec.name_count > 0 && spec.names[0] != nullptr && spec.name_lens[0] > 0)
    {
        column.name = std::string(reinterpret_cast<const char *>(spec.names[0]),
                                  static_cast<size_t>(spec.name_lens[0]));
    }
    column.index = spec.index;
    column.type = spec.type;
    column.nullable = spec.nullable != 0;
    columns.push_back(std::move(column));
}
return columns;
```

- [ ] **Step 7: Build and run**

Run: `cmake --build build --target excelreader_smoke && ./build/cpp/tests/excelreader_smoke` (adjust path/target name per the actual CMake layout)
Expected: PASS, including the new `test_parse_with_alias` and the existing single-name `make_field("Coluna1", &Row::Coluna1)` calls (unchanged source) still compiling and passing.

- [ ] **Step 8: Commit**

```bash
git add src/ExcelReader.Native/include/excelreader.hpp cpp/tests/smoke.cpp
git commit -m "feat(cpp): make_field accepts an ordered list of candidate column names

FieldBinding<Class, T, N> stores N candidate names (default N=1, source-
compatible with the existing single-name make_field); build_specs marshals
them into xl_column_spec's names/name_lens/name_count."
```

---

## Task 5: Rust binding

**Files:**
- Modify: `rust/excelreader/src/lib.rs:25,77-84`
- Modify: `rust/excelreader/src/workbook.rs:239-260,295-299,499-514`
- Modify: `rust/excelreader-derive/src/lib.rs:58-103`
- Test: `rust/excelreader-derive/src/lib.rs` (inline `#[cfg(test)]` module), `rust/excelreader/tests/parse_typed.rs`

**Interfaces:**
- Consumes: the `xl_column_spec` shape from Task 1, Step 3.
- Produces: `ColumnBinding<T>.names: &'static [&'static str]` (replaces `.name`), the `#[excel(name = "...", alias = "...")]` macro syntax.

- [ ] **Step 1: Write the failing derive-macro test**

Add to `rust/excelreader-derive/src/lib.rs`'s `#[cfg(test)] mod tests`:

```rust
#[test]
fn collects_aliases_in_declared_order_after_the_primary_name() {
    let output = expand_str(
        r#"
        struct Row {
            #[excel(name = "Nome", alias = "Nom", alias = "Name")]
            nome: String,
        }
        "#,
    )
    .expect("expand must succeed");

    let nome_pos = output.find("\"Nome\"").expect("primary name must appear");
    let nom_pos = output.find("\"Nom\"").expect("first alias must appear");
    let name_pos = output.rfind("\"Name\"").expect("second alias must appear");
    assert!(nome_pos < nom_pos, "primary name must precede its aliases");
    assert!(nom_pos < name_pos, "aliases must stay in declared order");
}
```

- [ ] **Step 2: Run to verify it fails**

Run: `cargo test -p excelreader-derive collects_aliases_in_declared_order`
Expected: FAIL — `#[excel(alias = ...)]` is currently rejected by `excel_name`'s `parse_nested_meta` closure ("unsupported #[excel(...)] key, expected `name`").

- [ ] **Step 3: Bump XL_ABI_VERSION and change XlColumnSpec**

In `rust/excelreader/src/lib.rs`, line 25:

```rust
pub const XL_ABI_VERSION: i32 = 2;
```

Replace lines 77-84:

```rust
#[repr(C)]
pub struct XlColumnSpec {
    pub names: *const *const u8,
    pub name_lens: *const i32,
    pub name_count: i32,
    pub index: i32,
    pub r#type: i32,
    pub nullable: i32,
}
```

- [ ] **Step 4: Update ColumnBinding and copy_inferred in workbook.rs**

Replace `ColumnBinding` (lines 295-299):

```rust
/// One field <-> column binding for `T`. Construct via `ExcelMapper::bindings()`.
pub struct ColumnBinding<T> {
    /// Candidate header names, in priority order — the first one present in the header row wins.
    pub names: &'static [&'static str],
    pub xl_type: i32,
    pub assign: fn(&mut T, &XlColumn, i64),
}
```

Replace `copy_inferred` (lines 239-260):

```rust
unsafe fn copy_inferred(schema: &XlInferredSchema) -> Vec<InferredColumn> {
    if schema.columns.is_null() || schema.column_count <= 0 {
        return Vec::new();
    }
    let specs = std::slice::from_raw_parts(schema.columns, schema.column_count as usize);
    specs
        .iter()
        .map(|spec| InferredColumn {
            // A guessed name is exactly name_lens[0] bytes with no NUL terminator, and name_count
            // is 0 whenever the column had no usable header cell.
            name: if spec.name_count <= 0 || spec.names.is_null() || spec.name_lens.is_null() {
                None
            } else {
                let name_ptr = *spec.names;
                let name_len = *spec.name_lens;
                if name_ptr.is_null() || name_len <= 0 {
                    None
                } else {
                    let bytes = std::slice::from_raw_parts(name_ptr, name_len as usize);
                    Some(String::from_utf8_lossy(bytes).into_owned())
                }
            },
            index: spec.index,
            column_type: spec.r#type,
            nullable: spec.nullable != 0,
        })
        .collect()
}
```

- [ ] **Step 5: Update parse_sheet's spec-building (lines 499-514)**

```rust
pub fn parse_sheet<T: ExcelMapper>(
    workbook: &mut Workbook,
    header_row: i32,
) -> Result<TableView<T>, Error> {
    let bindings = T::bindings();
    let name_ptrs: Vec<Vec<*const u8>> = bindings
        .iter()
        .map(|b| b.names.iter().map(|n| n.as_ptr()).collect())
        .collect();
    let name_lens: Vec<Vec<i32>> = bindings
        .iter()
        .map(|b| b.names.iter().map(|n| n.len() as i32).collect())
        .collect();
    let specs: Vec<XlColumnSpec> = bindings
        .iter()
        .enumerate()
        .map(|(i, b)| XlColumnSpec {
            names: name_ptrs[i].as_ptr(),
            name_lens: name_lens[i].as_ptr(),
            name_count: b.names.len() as i32,
            index: 0,
            r#type: b.xl_type,
            nullable: 1,
        })
        .collect();
```

(the rest of `parse_sheet` — building `table`, calling `xl_parse_typed`, checking `status` — is unchanged; `name_ptrs`/`name_lens` are local `Vec`s that outlive the call, same lifetime as the existing `specs` local).

- [ ] **Step 6: Update the derive macro**

Replace `excel_name` (lines 80-103 of `rust/excelreader-derive/src/lib.rs`) with `excel_names`:

```rust
fn excel_names(field: &Field) -> syn::Result<Vec<LitStr>> {
    let mut name = None;
    let mut aliases = Vec::new();
    for attr in &field.attrs {
        if !attr.path().is_ident("excel") {
            continue;
        }
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("name") {
                name = Some(meta.value()?.parse::<LitStr>()?);
                Ok(())
            } else if meta.path.is_ident("alias") {
                aliases.push(meta.value()?.parse::<LitStr>()?);
                Ok(())
            } else {
                Err(meta.error("unsupported #[excel(...)] key, expected `name` or `alias`"))
            }
        })?;
    }
    let Some(name) = name else {
        return Err(syn::Error::new(
            field.span(),
            "field is missing #[excel(name = \"...\")]",
        ));
    };
    let mut names = vec![name];
    names.extend(aliases);
    Ok(names)
}
```

Update `field_binding` (lines 58-78) to call it and emit a slice:

```rust
fn field_binding(field: &Field) -> syn::Result<proc_macro2::TokenStream> {
    let field_ident = field.ident.as_ref().expect("named_fields guarantees Some");
    let names = excel_names(field)?;
    let (inner_ty, is_option) = unwrap_option(&field.ty);
    let kind = FieldKind::from_type(inner_ty)?;
    let xl_type = kind.xl_type_tokens();
    let value = kind.value_tokens();
    let assign_value = if is_option {
        quote! { ::std::option::Option::Some(#value) }
    } else {
        value
    };

    Ok(quote! {
        ::excelreader::workbook::ColumnBinding {
            names: &[#(#names),*],
            xl_type: #xl_type,
            assign: |r, col, row| r.#field_ident = #assign_value,
        }
    })
}
```

- [ ] **Step 7: Run derive-crate tests**

Run: `cargo test -p excelreader-derive`
Expected: PASS, including the new alias test and every existing test (`generates_one_binding_per_field_with_inferred_types`, `errors_when_name_attribute_is_missing`, etc. — none of their assertions reference the removed `excel_name` function name, only the emitted token substrings, which are unchanged for the single-`name` case).

- [ ] **Step 8: Write and run the end-to-end alias test**

Add to `rust/excelreader/tests/parse_typed.rs`, after the existing struct definitions:

```rust
#[derive(Default, ExcelMapper)]
struct AliasRow {
    #[excel(name = "ThisColumnDoesNotExist", alias = "Coluna1")]
    coluna1: String,
}

#[test]
fn resolves_the_first_alias_present_in_the_header_row() {
    let mut workbook = open_fixture();
    let table = parse_sheet::<AliasRow>(&mut workbook, 1).expect("parse_sheet must succeed via alias");
    assert_eq!(table.len(), 100);
    let first = table.get(0).expect("row 0 is in bounds");
    assert_eq!(first.coluna1, "Valor1");
}
```

Run: `cargo test -p excelreader --test parse_typed`
Expected: PASS, including the new test and every existing one in the file.

- [ ] **Step 9: Commit**

```bash
git add rust/excelreader/src/lib.rs rust/excelreader/src/workbook.rs rust/excelreader-derive/src/lib.rs rust/excelreader/tests/parse_typed.rs
git commit -m "feat(rust): support #[excel(alias = \"...\")] for multi-name column resolution

ColumnBinding.names replaces .name (a &'static [&'static str] instead of
a single &'static str); the derive macro collects name plus every alias
in declared order. Bumps XL_ABI_VERSION to 2."
```

---

## Task 6: Python binding

**Files:**
- Modify: `python/src/excelreader/_native.py:29,70-106`
- Modify: `python/src/excelreader/types.py:79-90`
- Modify: `python/src/excelreader/reader.py:391-409`
- Test: `python/tests/test_native.py`, `python/tests/test_reader.py`

**Interfaces:**
- Consumes: the `xl_column_spec` shape from Task 1, Step 3.
- Produces: `ColumnSpec.name: str | Sequence[str] | None`, `_native.column_spec_by_names(names, type_, *, nullable=False)`.

- [ ] **Step 1: Write the failing test**

Add to `python/tests/test_native.py`, after `test_parse_typed_returns_typed_columns_by_name`:

```python
def test_parse_typed_resolves_the_first_alias_present_in_the_header_row(tmp_path):
    lib = _native.load_library()
    csv_file = tmp_path / "typed.csv"
    csv_file.write_text("name,qty\nwidget,3\ngadget,7\n", encoding="utf-8")
    path = str(csv_file).encode("utf-8")
    handle = ctypes.c_void_p()
    assert lib.xl_open_file(path, len(path), _native.XL_FORMAT_CSV, ctypes.byref(handle)) == _native.XL_OK

    specs = (_native.NativeColumnSpec * 2)(
        _native.column_spec_by_names(["does-not-exist", "name"], _native.XL_T_STRING),
        _native.column_spec_by_name("qty", _native.XL_T_I64),
    )
    table = _native.NativeTable()
    status = lib.xl_parse_typed(handle, specs, len(specs), 1, ctypes.byref(table))
    assert status == _native.XL_OK
    assert table.row_count == 2

    lib.xl_free_table(ctypes.byref(table))
    lib.xl_close(handle)
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest python/tests/test_native.py::test_parse_typed_resolves_the_first_alias_present_in_the_header_row -v`
Expected: FAIL — `_native` has no `column_spec_by_names` yet, and `NativeColumnSpec`/`XL_ABI_VERSION` don't match the C header's new layout/version, so `load_library()` itself would also start raising once the native lib is rebuilt from Task 1/2.

- [ ] **Step 3: Bump XL_ABI_VERSION and change the ctypes structs in _native.py**

Line 29:

```python
XL_ABI_VERSION = 2
```

Replace `NativeColumnSpec` and its helpers (lines 70-88):

```python
class NativeColumnSpec(ctypes.Structure):
    """Mirrors xl_column_spec. `names`/`name_lens`/`name_count` may be left NULL/NULL/0 to resolve by
    `index` instead. Build one with `column_spec_by_name`/`column_spec_by_names`, never directly —
    the `names`/`name_lens` pointers must stay alive as long as the struct is in use, and those
    helpers keep the backing buffers alive via ctypes' `_objects` mechanism."""

    _fields_ = [
        ("names", ctypes.POINTER(ctypes.POINTER(ctypes.c_uint8))),
        ("name_lens", ctypes.POINTER(ctypes.c_int32)),
        ("name_count", ctypes.c_int32),
        ("index", ctypes.c_int32),
        ("type", ctypes.c_int32),
        ("nullable", ctypes.c_int32),
    ]


def column_spec_by_names(names: Sequence[str], type_: int, *, nullable: bool = False) -> NativeColumnSpec:
    encoded = [name.encode("utf-8") for name in names]
    # One ctypes buffer per name (kept alive by `spec._objects` through the pointer array below),
    # plus the pointer array and length array themselves.
    buffers = [ctypes.create_string_buffer(e, len(e)) for e in encoded]
    name_ptrs = (ctypes.POINTER(ctypes.c_uint8) * len(buffers))(
        *(ctypes.cast(b, ctypes.POINTER(ctypes.c_uint8)) for b in buffers)
    )
    name_lens = (ctypes.c_int32 * len(encoded))(*(len(e) for e in encoded))
    spec = NativeColumnSpec(
        names=ctypes.cast(name_ptrs, ctypes.POINTER(ctypes.POINTER(ctypes.c_uint8))),
        name_lens=name_lens,
        name_count=len(encoded),
        index=0,
        type=type_,
        nullable=int(nullable),
    )
    # ctypes only auto-keeps-alive objects assigned directly to a field; `name_ptrs`/`name_lens`/
    # `buffers` were only cast/wrapped, so pin them explicitly on the returned struct.
    spec._name_storage = (buffers, name_ptrs, name_lens)  # type: ignore[attr-defined]
    return spec


def column_spec_by_name(name: str, type_: int, *, nullable: bool = False) -> NativeColumnSpec:
    return column_spec_by_names([name], type_, nullable=nullable)


def column_spec_by_index(index: int, type_: int, *, nullable: bool = False) -> NativeColumnSpec:
    return NativeColumnSpec(
        names=None, name_lens=None, name_count=0, index=index, type=type_, nullable=int(nullable)
    )
```

Add `from typing import Sequence` (or `from collections.abc import Sequence`, matching whatever this file already imports elsewhere — check its existing `TYPE_CHECKING` import block) near the top.

Replace `NativeInferredColumnSpec` (lines 91-106):

```python
class NativeInferredColumnSpec(ctypes.Structure):
    """Mirrors xl_column_spec's layout exactly, for the OUTPUT direction (xl_infer_schema).

    Always carries `name_count` 0 or 1 — inference never guesses more than one candidate name per
    column. Unlike `column_spec_by_name`'s buffers, `names[0]` here is a raw pointer with no
    guaranteed NUL terminator, so decode it with `ctypes.string_at(names[0], name_lens[0])`.
    """

    _fields_ = [
        ("names", ctypes.POINTER(ctypes.POINTER(ctypes.c_uint8))),
        ("name_lens", ctypes.POINTER(ctypes.c_int32)),
        ("name_count", ctypes.c_int32),
        ("index", ctypes.c_int32),
        ("type", ctypes.c_int32),
        ("nullable", ctypes.c_int32),
    ]
```

- [ ] **Step 4: Update reader.py's decode/build functions**

Replace `_decode_inferred_schema` (lines 391-397 of `python/src/excelreader/reader.py`):

```python
def _decode_inferred_schema(schema: _native.NativeInferredSchema) -> list[ColumnSpec]:
    specs: list[ColumnSpec] = []
    for index in range(schema.column_count):
        raw = schema.columns[index]
        name = None
        if raw.name_count > 0:
            name = ctypes.string_at(raw.names[0], raw.name_lens[0]).decode("utf-8")
        specs.append(ColumnSpec(ColumnType(raw.type), name=name, index=raw.index, nullable=bool(raw.nullable)))
    return specs
```

Replace `_build_specs` (lines 400-409):

```python
def _build_specs(schema: Sequence[ColumnSpec]) -> ctypes.Array:
    if not schema:
        raise ValueError("schema must name at least one column")
    specs = (_native.NativeColumnSpec * len(schema))()
    for index, spec in enumerate(schema):
        if spec.name is None:
            specs[index] = _native.column_spec_by_index(spec.index, int(spec.type), nullable=spec.nullable)
        elif isinstance(spec.name, str):
            specs[index] = _native.column_spec_by_name(spec.name, int(spec.type), nullable=spec.nullable)
        else:
            specs[index] = _native.column_spec_by_names(spec.name, int(spec.type), nullable=spec.nullable)
    return specs
```

- [ ] **Step 5: Widen ColumnSpec.name in types.py**

Replace lines 79-90 of `python/src/excelreader/types.py`:

```python
class ColumnSpec(NamedTuple):
    """One column to read, and the type to convert it to.

    Leave `name` as None to resolve the column by `index` instead. A `str` names exactly one header;
    a `Sequence[str]` names an ordered list of candidates, resolved first-match-wins against the
    header row — the same "try each in order, keep the first hit" rule the C# core's
    `[ExcelColumn(AllowMultiple = true)]` uses. `nullable` decides what a failed conversion means:
    False makes it an error that aborts the whole read, True records a null in the column's validity
    bitmap and keeps going.
    """

    type: ColumnType
    name: str | Sequence[str] | None = None
    index: int = 0
    nullable: bool = False
```

Check the top of `types.py` for its existing imports — add `from collections.abc import Sequence` (or reuse whatever it already imports) if `Sequence` isn't already in scope there.

- [ ] **Step 6: Run the Python test suite**

Run: `pytest python/tests -v`
Expected: PASS, including the new alias test, `test_native.py`'s `test_abi_version_matches_header` (now comparing `2 == 2`), and every existing `test_reader.py`/`test_writer.py` case (all pass a plain `str` name, unaffected).

- [ ] **Step 7: Add a widened-schema unit test**

Add to `python/tests/test_reader.py`, right after `test_parse_typed_resolves_columns_by_index_when_name_is_none` (line 300-305), reusing the same `typed_csv` fixture (defined at line 262, columns `name,qty,price,flag,day,clock,stamp`, `qty` values `[3, 7]`):

```python
def test_parse_typed_accepts_a_sequence_of_candidate_names(typed_csv):
    with open_workbook(typed_csv) as workbook:
        table = workbook.parse_typed([ColumnSpec(ColumnType.I64, name=["does-not-exist", "qty"])])

    assert list(table.columns[0]) == [3, 7]
```

Run: `pytest python/tests/test_reader.py -k candidate_names -v`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add python/src/excelreader/_native.py python/src/excelreader/types.py python/src/excelreader/reader.py python/tests/test_native.py python/tests/test_reader.py
git commit -m "feat(python): accept a sequence of candidate column names in ColumnSpec

ColumnSpec.name widens to str | Sequence[str] | None; _native gains
column_spec_by_names alongside the existing single-name helper. Bumps
XL_ABI_VERSION to 2 to match the native library."
```

---

## Out of scope (per spec)

- Any change to the C# core (`ExcelColumnAttribute`/`TypeMapper`) — already has this feature.
- A duplicate-header-claim analyzer for the native bindings (the C# source generator's `DuplicateHeaderDescriptor` has no native-binding equivalent here).
- Changing the header-matching rule itself (case-insensitive, trimmed) — unchanged, just applied per candidate.
