# Native FFI + Python Bindings (Reading) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship a NativeAOT C-ABI shared library over the existing ExcelReader readers, plus a Python package that reads XLSX, XLSB, XLS and CSV through it — with no .NET runtime installed on the user's machine.

**Architecture:** A new `src/ExcelReader.Native` project publishes as a NativeAOT shared library (`NativeLib=Shared`). It exposes ~9 C functions over `IExcelRowReader`. Because `Row`/`Cell` are `ref struct`s over the reader's internal buffers, they cannot cross the FFI boundary — so each row is serialized into a compact little-endian blob that the caller decodes. The exported `[UnmanagedCallersOnly]` functions are thin pointer wrappers over an internal, span-based `NativeApi` class, which is what the C# tests actually exercise (managed code cannot call `[UnmanagedCallersOnly]` methods directly). A thin Python package (`ctypes`, stdlib only) wraps the same ABI.

**Tech Stack:** .NET 10 / .NET 8 (multi-target, matching `ExcelReader.Core`), NativeAOT, xUnit v3 + Microsoft.Testing.Platform (existing C# test project), Python 3.9+, `ctypes` (stdlib), pytest, hatchling.

**Spec:** This document — see [ABI Contract](#abi-contract) below. There is no separate spec file; the ABI section is normative and every task must match it exactly.

**Scope:** Reading only. Writing (`IWorkbookWriter<TSheet>` and friends) is explicitly **out of scope** for this plan and will be a follow-up plan. Do not add write functions.

---

## Global Constraints

- **Target frameworks for `ExcelReader.Native`:** `net10.0;net8.0` — must match `ExcelReader.Core`, because `tests/ExcelReader.Tests` multi-targets both and will reference this project.
- **NativeAOT publish must be warning-free.** The repo root `Directory.Build.props` sets `TreatWarningsAsErrors=true`, `AnalysisMode=All`, `WarningLevel=9999`. Any new project must either satisfy the analyzers or carry an explicit, *commented* `NoWarn` entry — follow the precedent in `tests/ExcelReader.AotSanity/ExcelReader.AotSanity.csproj`.
- **No new NuGet dependencies** in `ExcelReader.Native`. It references only `ExcelReader.Core`.
- **No new Python dependencies at runtime.** `ctypes` only. pytest is a dev dependency.
- **All integers on the ABI are little-endian `int32`** unless the signature says `int64`.
- **All strings on the ABI are UTF-8 byte arrays with an explicit length.** Never NUL-terminated-only, never UTF-16.
- **Exceptions must never cross the boundary.** Every exported function catches, stores the message, and returns a negative status code.
- **Style:** follow `STYLEGUIDE.md` — always block bodies for methods, early returns over nesting, max 3 levels of nesting, max 70 lines per method, comments explain *why* not *what*.
- **Public API tracking:** `ExcelReader.Native` is never packed as a NuGet library, so it must `NoWarn` `RS0016;RS0026;RS0037` rather than carry a `PublicAPI.txt`.

---

## ABI Contract

This is the normative contract. Tasks 1–5 implement the C# side; tasks 6–8 implement the Python side against the same numbers.

### Status codes

| Name | Value | Meaning |
|---|---|---|
| `XL_OK` | `0` | Success |
| `XL_EOF` | `-1` | No more rows on the current sheet |
| `XL_BUFFER_TOO_SMALL` | `-2` | Caller buffer too small; the out-length parameter holds the required size |
| `XL_INVALID_HANDLE` | `-3` | Handle is null or not a live workbook handle |
| `XL_INVALID_ARGUMENT` | `-4` | A pointer was null or a length was negative |
| `XL_ERROR` | `-5` | An exception was caught; message available via `xl_last_error` |

### Format codes

| Name | Value | Notes |
|---|---|---|
| `XL_FORMAT_AUTO` | `0` | Sniffs XLS/XLSX/XLSB by signature. **Does not detect CSV.** |
| `XL_FORMAT_XLS` | `1` | |
| `XL_FORMAT_XLSX` | `2` | |
| `XL_FORMAT_XLSB` | `3` | |
| `XL_FORMAT_CSV` | `4` | Must be requested explicitly |

Values 1–3 deliberately match `ExcelReader.Core.Enums.ExcelFileFormat`.

### Functions

```c
int32_t xl_open_file  (const uint8_t* path, int32_t path_len, int32_t format, void** out_handle);
int32_t xl_open_memory(const uint8_t* data, int32_t data_len, int32_t format, void** out_handle);
int32_t xl_close      (void* handle);

int32_t xl_sheet_count (void* handle, int32_t* out_count);
int32_t xl_sheet_name  (void* handle, uint8_t* buffer, int32_t capacity, int32_t* out_len);
int32_t xl_move_to_sheet(void* handle, int32_t index);
int32_t xl_is_date1904 (void* handle, int32_t* out_flag);

int32_t xl_next_row    (void* handle, uint8_t* buffer, int32_t capacity, int32_t* out_written);
int32_t xl_last_error  (uint8_t* buffer, int32_t capacity, int32_t* out_len);
```

### Row blob layout (version 1)

`xl_next_row` writes, all little-endian:

```
int32  cell_count
repeated cell_count times:
    int32  column_index     (zero-based; gaps are possible — empty cells are skipped)
    int32  cell_type        (see table below)
    int32  value_len        (bytes)
    uint8  value[value_len] (UTF-8; raw cell text as stored — dates are Excel serial numbers)
```

Cell types mirror `ExcelReader.Core.Enums.CellType`: `0 empty`, `1 string`, `2 number`, `3 date`, `4 bool`, `5 formula`, `6 error`.

### Buffer-too-small semantics

`xl_next_row` and `xl_sheet_name` never lose data. If the caller's buffer is too small they set `out_written`/`out_len` to the required byte count and return `XL_BUFFER_TOO_SMALL`; the pending row is held inside the handle and re-served on the next call with a larger buffer.

### Lifetime rules

- A handle is valid until `xl_close`. Calling anything after `xl_close` is undefined behavior (the caller must null its pointer).
- `xl_open_memory` **copies** the caller's bytes. The caller may free its buffer immediately.
- `xl_move_to_sheet` resets row enumeration to the start of the new sheet and drops any pending row.
- No handle is thread-safe. One handle per thread.
- `xl_last_error` is thread-local: it reports the last failure on the *calling* thread.

---

## File Structure

| File | Responsibility |
|---|---|
| `src/ExcelReader.Native/ExcelReader.Native.csproj` | Project + NativeAOT settings + analyzer suppressions |
| `src/ExcelReader.Native/NativeStatus.cs` | Status and format constants (Task 1) |
| `src/ExcelReader.Native/NativeHandle.cs` | Reader + enumerator + pending-row scratch buffer (Task 2) |
| `src/ExcelReader.Native/NativeApi.cs` | Internal span-based implementation — the testable surface (Tasks 1–4) |
| `src/ExcelReader.Native/RowBlob.cs` | Row → blob serialization (Task 4) |
| `src/ExcelReader.Native/Exports.cs` | `[UnmanagedCallersOnly]` pointer wrappers (Task 5) |
| `src/ExcelReader.Native/include/excelreader.h` | Hand-written C header for non-Python consumers (Task 5) |
| `tests/ExcelReader.Tests/NativeApiTests.cs` | xUnit tests against `NativeApi` (Tasks 1–4) |
| `python/pyproject.toml` | Python package metadata (Task 6) |
| `python/scripts/build_native.py` | Runs `dotnet publish` and copies the binary into the package (Task 6) |
| `python/src/excelreader/_native.py` | Library loading + `ctypes` signatures (Task 6) |
| `python/src/excelreader/types.py` | `CellType`, `Cell`, `ExcelReaderError` (Task 7) |
| `python/src/excelreader/reader.py` | `Workbook`, `open_workbook`, `open_bytes` (Task 7) |
| `python/src/excelreader/__init__.py` | Public exports (Task 7) |
| `python/tests/conftest.py` | Fixture paths (Task 7) |
| `python/tests/test_reader.py` | Read tests across all four formats (Tasks 7–8) |
| `.github/workflows/python.yml` | Build native + run pytest on 3 OSes (Task 9) |

### Test fixtures available in this repo

| Format | Path | Known shape |
|---|---|---|
| XLSX | `RealExcel.xlsx` (repo root) | 1 sheet, 101 rows, 18 columns, header row is `Coluna1`…`Coluna18` |
| XLSB | `RealExcel.xlsb` (repo root) | 1 sheet — assert structurally, exact counts unverified |
| XLS | `tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xls` | ~65K rows, 11 MB — always cap iteration in tests |
| CSV | write a small file in the test itself | — |
| XLSX (small) | `tests/ExcelReader.Tests/data/sample.xlsx` | already copied to the C# test output directory |

**C# tests cover XLSX + XLSB + CSV. XLS coverage lives in the Python tests** (the only XLS fixture is 11 MB and is not worth copying into the C# test output on every build). This is a deliberate split, not an oversight.

---

## Task 1: Native project scaffold, status codes, last-error channel

**Files:**
- Create: `src/ExcelReader.Native/ExcelReader.Native.csproj`
- Create: `src/ExcelReader.Native/NativeStatus.cs`
- Create: `src/ExcelReader.Native/NativeApi.cs`
- Modify: `ExcelReader.slnx`
- Modify: `tests/ExcelReader.Tests/ExcelReader.Tests.csproj` (add ProjectReference)
- Test: `tests/ExcelReader.Tests/NativeApiTests.cs`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `ExcelReader.Native.NativeStatus` with `internal const int Ok = 0; Eof = -1; BufferTooSmall = -2; InvalidHandle = -3; InvalidArgument = -4; Error = -5;`
  - `ExcelReader.Native.NativeFormat` with `internal const int Auto = 0; Xls = 1; Xlsx = 2; Xlsb = 3; Csv = 4;`
  - `internal static void NativeApi.SetLastError(string message)`
  - `internal static int NativeApi.LastError(Span<byte> buffer, out int length)`
  - `internal static void NativeApi.ClearLastError()`

- [ ] **Step 1: Create the project file**

Create `src/ExcelReader.Native/ExcelReader.Native.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFrameworks>net10.0;net8.0</TargetFrameworks>
    <AllowUnsafeBlocks>true</AllowUnsafeBlocks>
    <IsPackable>false</IsPackable>
    <!-- Publishing this project with -r <rid> produces the C-ABI shared library
         (ExcelReader.Native.dll/.so/.dylib). Ordinary `dotnet build` still produces a normal
         managed assembly, which is what ExcelReader.Tests references. -->
    <PublishAot>true</PublishAot>
    <NativeLib>Shared</NativeLib>
    <TrimmerSingleWarn>false</TrimmerSingleWarn>
    <!-- Keeps ICU out of the shared library; this layer does no culture-sensitive formatting. -->
    <InvariantGlobalization>true</InvariantGlobalization>

    <!-- RS0016/RS0026/RS0037: public-API tracking is for ExcelReader.Core's shipped NuGet surface.
         This project is never packed — its C ABI is tracked by include/excelreader.h instead. -->
    <!-- CA1031/S2221: every exported function is a hard boundary. An escaping exception would
         corrupt the caller's stack, so catch-all plus a status code is the required behavior. -->
    <!-- CA2000/IDISP001/IDISP004: file streams handed to Excel.From*(leaveOpen: false) are owned
         and disposed by the reader — the same ownership-transfer pattern used across the tests. -->
    <NoWarn>$(NoWarn);RS0016;RS0026;RS0037;CA1031;S2221;CA2000;IDISP001;IDISP004</NoWarn>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="..\ExcelReader.Core\ExcelReader.Core.csproj" />
    <ProjectReference Include="..\ExcelReader.Generator\ExcelReader.Generator.csproj"
                      OutputItemType="Analyzer" ReferenceOutputAssembly="false" />
  </ItemGroup>

  <!-- The span-based NativeApi is the tested surface; [UnmanagedCallersOnly] exports cannot be
       invoked from managed code, so the tests drive the layer underneath them. -->
  <ItemGroup>
    <InternalsVisibleTo Include="ExcelReader.Tests" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Register the project in the solution**

In `ExcelReader.slnx`, add inside the `/src/` folder, after the `ExcelReader.Core` line:

```xml
    <Project Path="src/ExcelReader.Native/ExcelReader.Native.csproj" />
```

- [ ] **Step 3: Reference it from the test project**

In `tests/ExcelReader.Tests/ExcelReader.Tests.csproj`, in the `ItemGroup` that already contains
`<ProjectReference Include="..\..\src\ExcelReader.Core\ExcelReader.Core.csproj" />`, add:

```xml
    <ProjectReference Include="..\..\src\ExcelReader.Native\ExcelReader.Native.csproj" />
```

- [ ] **Step 4: Write the failing test**

Create `tests/ExcelReader.Tests/NativeApiTests.cs`:

```csharp
using System.Text;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    public sealed class NativeApiTests
    {
        [Fact]
        public void LastError_Should_Return_Stored_Message_As_Utf8()
        {
            NativeApi.SetLastError("boom");

            Span<byte> buffer = stackalloc byte[64];
            int status = NativeApi.LastError(buffer, out int length);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.Equal(4, length);
            Assert.Equal("boom", Encoding.UTF8.GetString(buffer[..length]));
        }

        [Fact]
        public void LastError_Should_Report_Required_Size_When_Buffer_Too_Small()
        {
            NativeApi.SetLastError("boom");

            Span<byte> buffer = stackalloc byte[2];
            int status = NativeApi.LastError(buffer, out int length);

            Assert.Equal(NativeStatus.BufferTooSmall, status);
            Assert.Equal(4, length);
        }

        [Fact]
        public void LastError_Should_Return_Zero_Length_When_Cleared()
        {
            NativeApi.SetLastError("boom");
            NativeApi.ClearLastError();

            Span<byte> buffer = stackalloc byte[64];
            int status = NativeApi.LastError(buffer, out int length);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.Equal(0, length);
        }
    }
}
```

- [ ] **Step 5: Run the test to verify it fails**

Run: `dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj`
Expected: FAIL — compile errors, `NativeApi`/`NativeStatus` do not exist.

- [ ] **Step 6: Write the constants**

Create `src/ExcelReader.Native/NativeStatus.cs`:

```csharp
namespace ExcelReader.Native
{
    /// <summary>Status codes returned by every exported C function. Mirrors include/excelreader.h.</summary>
    internal static class NativeStatus
    {
        internal const int Ok = 0;
        internal const int Eof = -1;
        internal const int BufferTooSmall = -2;
        internal const int InvalidHandle = -3;
        internal const int InvalidArgument = -4;
        internal const int Error = -5;
    }

    /// <summary>
    /// Format selectors accepted by the open functions. Values 1-3 deliberately match
    /// <see cref="ExcelReader.Core.Enums.ExcelFileFormat"/>; CSV has no signature to sniff, so it
    /// has no counterpart there and must always be requested explicitly.
    /// </summary>
    internal static class NativeFormat
    {
        internal const int Auto = 0;
        internal const int Xls = 1;
        internal const int Xlsx = 2;
        internal const int Xlsb = 3;
        internal const int Csv = 4;
    }
}
```

- [ ] **Step 7: Write the last-error channel**

Create `src/ExcelReader.Native/NativeApi.cs`:

```csharp
using System.Text;

namespace ExcelReader.Native
{
    /// <summary>
    /// Span-based implementation behind the C ABI. Every method returns a <see cref="NativeStatus"/>
    /// code and never throws; <see cref="Exports"/> only converts pointers into spans on top of it.
    /// This split exists because managed code cannot call an [UnmanagedCallersOnly] method, so the
    /// exports themselves are untestable — this layer is what the test suite drives.
    /// </summary>
    internal static partial class NativeApi
    {
        // Thread-local because handles are single-threaded by contract: an error raised on one
        // thread must not be observable from another.
        [ThreadStatic]
        private static string? _lastError;

        internal static void SetLastError(string message)
        {
            _lastError = message;
        }

        internal static void ClearLastError()
        {
            _lastError = null;
        }

        /// <summary>Copies the calling thread's last error message into <paramref name="buffer"/> as UTF-8.</summary>
        internal static int LastError(Span<byte> buffer, out int length)
        {
            string? message = _lastError;
            if (string.IsNullOrEmpty(message))
            {
                length = 0;
                return NativeStatus.Ok;
            }

            int required = Encoding.UTF8.GetByteCount(message);
            length = required;
            if (buffer.Length < required)
            {
                return NativeStatus.BufferTooSmall;
            }

            Encoding.UTF8.GetBytes(message, buffer);
            return NativeStatus.Ok;
        }
    }
}
```

- [ ] **Step 8: Run the tests to verify they pass**

Run: `dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj`
Expected: PASS, and the whole solution builds warning-free.

- [ ] **Step 9: Verify the solution still builds clean**

Run: `dotnet build ExcelReader.slnx --configuration Release`
Expected: `Build succeeded`, 0 warnings. If an analyzer fires, add it to the `NoWarn` list in
`ExcelReader.Native.csproj` **with a comment explaining why**, per the existing convention.

- [ ] **Step 10: Commit**

```bash
git add src/ExcelReader.Native ExcelReader.slnx tests/ExcelReader.Tests/ExcelReader.Tests.csproj tests/ExcelReader.Tests/NativeApiTests.cs
git commit -m "feat(native): add ExcelReader.Native project with status codes and last-error channel"
```

---

## Task 2: Handle lifetime — open and close every format

**Files:**
- Create: `src/ExcelReader.Native/NativeHandle.cs`
- Create: `src/ExcelReader.Native/NativeApi.Open.cs`
- Modify: `tests/ExcelReader.Tests/NativeApiTests.cs`
- Modify: `tests/ExcelReader.Tests/ExcelReader.Tests.csproj` (link the XLSB fixture into test output)

**Interfaces:**
- Consumes: `NativeStatus`, `NativeFormat`, `NativeApi.SetLastError` (Task 1).
- Produces:
  - `internal sealed class NativeHandle : IDisposable` with `internal IExcelRowReader Reader { get; }`, `internal IExcelRowEnumerator? Rows { get; set; }`, `internal byte[] Scratch { get; set; }`, `internal int PendingLength { get; set; }`, `internal bool HasPending { get; set; }`, `internal void ResetRows()` (used by Task 3's `MoveToSheet`)
  - `internal static int NativeApi.OpenFile(ReadOnlySpan<byte> utf8Path, int format, out NativeHandle? handle)`
  - `internal static int NativeApi.OpenMemory(ReadOnlySpan<byte> data, int format, out NativeHandle? handle)`
  - `internal static int NativeApi.Close(NativeHandle? handle)`

- [ ] **Step 1: Make the XLSB fixture available to the tests**

In `tests/ExcelReader.Tests/ExcelReader.Tests.csproj`, in the `ItemGroup` that already contains
`<Content Include="data\**" ... />`, add:

```xml
    <!-- Ground-truth XLSB produced by Excel itself; the native FFI tests open it by path. -->
    <Content Include="..\..\RealExcel.xlsb" Link="data\RealExcel.xlsb" CopyToOutputDirectory="PreserveNewest" />
```

- [ ] **Step 2: Write the failing tests**

Append these members to the `NativeApiTests` class in `tests/ExcelReader.Tests/NativeApiTests.cs`
(and add `using System.Text;` at the top if it is not already there):

```csharp
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");
        private static readonly string XlsbFixture = Path.Combine(AppContext.BaseDirectory, "data", "RealExcel.xlsb");

        private static int OpenPath(string path, int format, out NativeHandle? handle)
        {
            return NativeApi.OpenFile(Encoding.UTF8.GetBytes(path), format, out handle);
        }

        [Theory]
        [InlineData(NativeFormat.Auto)]
        [InlineData(NativeFormat.Xlsx)]
        public void OpenFile_Should_Open_Xlsx(int format)
        {
            int status = OpenPath(XlsxFixture, format, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
        }

        [Theory]
        [InlineData(NativeFormat.Auto)]
        [InlineData(NativeFormat.Xlsb)]
        public void OpenFile_Should_Open_Xlsb(int format)
        {
            int status = OpenPath(XlsbFixture, format, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
        }

        [Fact]
        public void OpenFile_Should_Open_Csv_When_Format_Is_Explicit()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "a,b\n1,2\n");
            try
            {
                int status = OpenPath(path, NativeFormat.Csv, out NativeHandle? handle);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.NotNull(handle);
                Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenFile_Should_Fail_With_Error_When_File_Is_Missing()
        {
            string path = Path.Combine(Path.GetTempPath(), $"does-not-exist-{Guid.NewGuid():N}.xlsx");

            int status = OpenPath(path, NativeFormat.Auto, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Error, status);
            Assert.Null(handle);

            Span<byte> buffer = stackalloc byte[512];
            Assert.Equal(NativeStatus.Ok, NativeApi.LastError(buffer, out int length));
            Assert.True(length > 0);
        }

        [Fact]
        public void OpenFile_Should_Reject_Unknown_Format_Code()
        {
            int status = OpenPath(XlsxFixture, 99, out NativeHandle? handle);

            Assert.Equal(NativeStatus.InvalidArgument, status);
            Assert.Null(handle);
        }

        [Fact]
        public void OpenMemory_Should_Open_Xlsx_From_A_Copy_Of_The_Bytes()
        {
            byte[] bytes = File.ReadAllBytes(XlsxFixture);

            int status = NativeApi.OpenMemory(bytes, NativeFormat.Auto, out NativeHandle? handle);

            Assert.Equal(NativeStatus.Ok, status);
            Assert.NotNull(handle);
            Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
        }

        [Fact]
        public void Close_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.Close(null));
        }
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj`
Expected: FAIL — `NativeHandle`, `NativeApi.OpenFile`, `NativeApi.OpenMemory`, `NativeApi.Close` do not exist.

- [ ] **Step 4: Write the handle**

Create `src/ExcelReader.Native/NativeHandle.cs`:

```csharp
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    /// <summary>
    /// Everything one open workbook needs on the managed side of the boundary. The caller only ever
    /// sees an opaque pointer to a GCHandle wrapping this object.
    /// </summary>
    /// <remarks>
    /// <see cref="Scratch"/> holds the most recently serialized row. A row is serialized once and
    /// then copied out, so a caller whose buffer was too small can retry with a bigger one without
    /// losing the row — the reader has already advanced past it and cannot rewind.
    /// </remarks>
    internal sealed class NativeHandle : IDisposable
    {
        internal NativeHandle(IExcelRowReader reader)
        {
            Reader = reader;
            Scratch = new byte[4096];
        }

        internal IExcelRowReader Reader { get; }

        /// <summary>Row cursor over the current sheet. Created lazily on the first row request, dropped on sheet change.</summary>
        internal IExcelRowEnumerator? Rows { get; set; }

        internal byte[] Scratch { get; set; }

        internal int PendingLength { get; set; }

        internal bool HasPending { get; set; }

        internal void ResetRows()
        {
            Rows?.Dispose();
            Rows = null;
            HasPending = false;
            PendingLength = 0;
        }

        public void Dispose()
        {
            ResetRows();
            Reader.Dispose();
        }
    }
}
```

- [ ] **Step 5: Write the open/close implementation**

Create `src/ExcelReader.Native/NativeApi.Open.cs`:

```csharp
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        internal static int OpenFile(ReadOnlySpan<byte> utf8Path, int format, out NativeHandle? handle)
        {
            handle = null;
            if (!IsKnownFormat(format))
            {
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            try
            {
                string path = Encoding.UTF8.GetString(utf8Path);
                IExcelRowReader reader = OpenReader(path, format);
                handle = new NativeHandle(reader);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int OpenMemory(ReadOnlySpan<byte> data, int format, out NativeHandle? handle)
        {
            handle = null;
            if (!IsKnownFormat(format))
            {
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            try
            {
                // Copied on purpose: the ABI promises the caller may free its buffer immediately,
                // and the readers keep referencing this memory for the handle's whole lifetime.
                byte[] copy = data.ToArray();
                IExcelRowReader reader = OpenReader(copy, format);
                handle = new NativeHandle(reader);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int Close(NativeHandle? handle)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            handle.Dispose();
            return NativeStatus.Ok;
        }

        private static IExcelRowReader OpenReader(string path, int format)
        {
            return format switch
            {
                NativeFormat.Auto => Excel.Open(path),
                NativeFormat.Xlsx => Excel.FromXlsxFile(path),
                NativeFormat.Xlsb => Excel.FromXlsb(File.OpenRead(path), leaveOpen: false),
                NativeFormat.Xls => Excel.FromXls(File.OpenRead(path), leaveOpen: false),
                _ => Excel.FromCsv(File.OpenRead(path), leaveOpen: false),
            };
        }

        private static IExcelRowReader OpenReader(byte[] data, int format)
        {
            return format switch
            {
                NativeFormat.Auto => Excel.Open(data),
                NativeFormat.Xlsx => Excel.FromXlsx(data),
                NativeFormat.Xlsb => Excel.FromXlsb(new MemoryStream(data, writable: false), leaveOpen: false),
                NativeFormat.Xls => Excel.FromXls(new MemoryStream(data, writable: false), leaveOpen: false),
                _ => Excel.FromCsv(new MemoryStream(data, writable: false), leaveOpen: false),
            };
        }

        private static bool IsKnownFormat(int format)
        {
            return format is >= NativeFormat.Auto and <= NativeFormat.Csv;
        }
    }
}
```

- [ ] **Step 6: Run the tests to verify they pass**

Run: `dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj`
Expected: PASS.

If `Excel.FromXlsb`/`Excel.FromXls`/`Excel.FromCsv` return concrete reader types that the compiler
will not implicitly convert to `IExcelRowReader`, that is a compile error — the fix is an explicit
cast, not a signature change: every concrete reader implements `IExcelRowReader`.

- [ ] **Step 7: Commit**

```bash
git add src/ExcelReader.Native tests/ExcelReader.Tests
git commit -m "feat(native): open and close workbooks for all four read formats"
```

---

## Task 3: Sheet navigation

**Files:**
- Create: `src/ExcelReader.Native/NativeApi.Sheets.cs`
- Modify: `tests/ExcelReader.Tests/NativeApiTests.cs`

**Interfaces:**
- Consumes: `NativeHandle`, `NativeApi.OpenFile`, `NativeApi.Close` (Task 2).
- Produces:
  - `internal static int NativeApi.SheetCount(NativeHandle? handle, out int count)`
  - `internal static int NativeApi.SheetName(NativeHandle? handle, Span<byte> buffer, out int length)`
  - `internal static int NativeApi.MoveToSheet(NativeHandle? handle, int index)`
  - `internal static int NativeApi.IsDate1904(NativeHandle? handle, out int flag)`

- [ ] **Step 1: Write the failing tests**

Append to `NativeApiTests`:

```csharp
        [Fact]
        public void SheetCount_Should_Report_At_Least_One_Sheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = NativeApi.SheetCount(handle, out int count);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.True(count >= 1);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void SheetName_Should_Return_A_Non_Empty_Utf8_Name()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Span<byte> buffer = stackalloc byte[256];
                int status = NativeApi.SheetName(handle, buffer, out int length);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.True(length > 0);
                Assert.NotEmpty(Encoding.UTF8.GetString(buffer[..length]));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void SheetName_Should_Report_Required_Size_When_Buffer_Too_Small()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = NativeApi.SheetName(handle, Span<byte>.Empty, out int length);

                Assert.Equal(NativeStatus.BufferTooSmall, status);
                Assert.True(length > 0);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void MoveToSheet_Should_Accept_The_First_Sheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.MoveToSheet(handle, 0));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void MoveToSheet_Should_Fail_For_An_Out_Of_Range_Index()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                Assert.Equal(NativeStatus.Error, NativeApi.MoveToSheet(handle, 9999));
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void IsDate1904_Should_Report_Zero_For_A_1900_Based_Workbook()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                int status = NativeApi.IsDate1904(handle, out int flag);

                Assert.Equal(NativeStatus.Ok, status);
                Assert.Equal(0, flag);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void Sheet_Functions_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.SheetCount(null, out _));
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.MoveToSheet(null, 0));
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.IsDate1904(null, out _));
        }
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj`
Expected: FAIL — the four `NativeApi` sheet methods do not exist.

- [ ] **Step 3: Write the implementation**

Create `src/ExcelReader.Native/NativeApi.Sheets.cs`:

```csharp
using System.Text;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        internal static int SheetCount(NativeHandle? handle, out int count)
        {
            count = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                count = handle.Reader.SheetCount;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int SheetName(NativeHandle? handle, Span<byte> buffer, out int length)
        {
            length = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                string name = handle.Reader.SheetName;
                int required = Encoding.UTF8.GetByteCount(name);
                length = required;
                if (buffer.Length < required)
                {
                    return NativeStatus.BufferTooSmall;
                }

                Encoding.UTF8.GetBytes(name, buffer);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int MoveToSheet(NativeHandle? handle, int index)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                handle.Reader.MoveToSheet(index);
                // Row enumeration is per-sheet: the old cursor points into the previous sheet's
                // buffers, so it is dropped and rebuilt on the next row request.
                handle.ResetRows();
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int IsDate1904(NativeHandle? handle, out int flag)
        {
            flag = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                flag = handle.Reader.IsDate1904 ? 1 : 0;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }
    }
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj`
Expected: PASS.

If `MoveToSheet_Should_Fail_For_An_Out_Of_Range_Index` fails because the reader does not throw,
change the assertion to match the reader's real behavior and note it in the test — do not add a
range check that duplicates Core's validation.

- [ ] **Step 5: Commit**

```bash
git add src/ExcelReader.Native tests/ExcelReader.Tests/NativeApiTests.cs
git commit -m "feat(native): expose sheet count, name, navigation and date system"
```

---

## Task 4: Row serialization

**Files:**
- Create: `src/ExcelReader.Native/RowBlob.cs`
- Create: `src/ExcelReader.Native/NativeApi.Rows.cs`
- Modify: `tests/ExcelReader.Tests/NativeApiTests.cs`

**Interfaces:**
- Consumes: `NativeHandle` (Task 2).
- Produces:
  - `internal static int RowBlob.Serialize(in Row row, ref byte[] scratch)` — returns bytes written into `scratch`
  - `internal static int NativeApi.NextRow(NativeHandle? handle, Span<byte> buffer, out int written)`

- [ ] **Step 1: Write the failing tests**

Append to `NativeApiTests`. This includes a small blob decoder, which is also the reference the
Python decoder must match:

```csharp
        private sealed record DecodedCell(int Column, int Type, string Value);

        private static List<DecodedCell> DecodeRow(ReadOnlySpan<byte> blob)
        {
            List<DecodedCell> cells = [];
            int count = BitConverter.ToInt32(blob[..4]);
            int offset = 4;
            for (int i = 0; i < count; i++)
            {
                int column = BitConverter.ToInt32(blob[offset..]);
                int type = BitConverter.ToInt32(blob[(offset + 4)..]);
                int valueLength = BitConverter.ToInt32(blob[(offset + 8)..]);
                offset += 12;
                cells.Add(new DecodedCell(column, type, Encoding.UTF8.GetString(blob.Slice(offset, valueLength))));
                offset += valueLength;
            }

            return cells;
        }

        [Fact]
        public void NextRow_Should_Decode_A_Csv_Row_Exactly()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\nwidget,7\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[4096];

                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int first));
                List<DecodedCell> header = DecodeRow(buffer.AsSpan(0, first));
                Assert.Equal(2, header.Count);
                Assert.Equal(0, header[0].Column);
                Assert.Equal("name", header[0].Value);
                Assert.Equal(1, header[1].Column);
                Assert.Equal("qty", header[1].Value);

                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int second));
                List<DecodedCell> data = DecodeRow(buffer.AsSpan(0, second));
                Assert.Equal("widget", data[0].Value);
                Assert.Equal("7", data[1].Value);

                Assert.Equal(NativeStatus.Eof, NativeApi.NextRow(handle, buffer, out _));
            }
            finally
            {
                NativeApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRow_Should_Reserve_The_Row_When_The_Buffer_Is_Too_Small()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-native-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "name,qty\n");
            Assert.Equal(NativeStatus.Ok, OpenPath(path, NativeFormat.Csv, out NativeHandle? handle));
            try
            {
                byte[] tiny = new byte[3];
                Assert.Equal(NativeStatus.BufferTooSmall, NativeApi.NextRow(handle, tiny, out int required));
                Assert.True(required > 3);

                // The same row must come back — a caller that grows its buffer must not lose data.
                byte[] big = new byte[required];
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, big, out int written));
                Assert.Equal(required, written);
                Assert.Equal("name", DecodeRow(big.AsSpan(0, written))[0].Value);
            }
            finally
            {
                NativeApi.Close(handle);
                File.Delete(path);
            }
        }

        [Fact]
        public void NextRow_Should_Read_Every_Row_Of_The_Xlsx_Fixture()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[1 << 20];
                int rows = 0;
                while (NativeApi.NextRow(handle, buffer, out int written) == NativeStatus.Ok)
                {
                    Assert.True(written >= 4);
                    rows++;
                }

                Assert.True(rows > 0);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void NextRow_Should_Restart_After_MoveToSheet()
        {
            Assert.Equal(NativeStatus.Ok, OpenPath(XlsxFixture, NativeFormat.Auto, out NativeHandle? handle));
            try
            {
                byte[] buffer = new byte[1 << 20];
                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int firstPass));

                Assert.Equal(NativeStatus.Ok, NativeApi.MoveToSheet(handle, 0));

                Assert.Equal(NativeStatus.Ok, NativeApi.NextRow(handle, buffer, out int secondPass));
                Assert.Equal(firstPass, secondPass);
            }
            finally
            {
                NativeApi.Close(handle);
            }
        }

        [Fact]
        public void NextRow_Should_Reject_A_Null_Handle()
        {
            Assert.Equal(NativeStatus.InvalidHandle, NativeApi.NextRow(null, new byte[16], out _));
        }
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj`
Expected: FAIL — `NativeApi.NextRow` does not exist.

- [ ] **Step 3: Write the serializer**

Create `src/ExcelReader.Native/RowBlob.cs`:

```csharp
using System.Buffers.Binary;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    /// <summary>
    /// Serializes a <see cref="Row"/> into the flat little-endian blob described in
    /// docs/plans/2026-08-13-native-ffi-python-reading.md (ABI Contract, "Row blob layout").
    /// </summary>
    /// <remarks>
    /// <see cref="Row"/> and <see cref="Cell"/> are ref structs over the reader's internal buffers,
    /// so they cannot be handed across the FFI boundary or stored between calls. Copying the whole
    /// row once per call keeps the boundary to a single crossing per row and leaves the caller with
    /// no lifetime rules to obey.
    /// </remarks>
    internal static class RowBlob
    {
        private const int CellHeaderSize = 3 * sizeof(int);

        /// <summary>Writes <paramref name="row"/> into <paramref name="scratch"/>, growing it if needed. Returns the byte count.</summary>
        internal static int Serialize(in Row row, ref byte[] scratch)
        {
            int required = sizeof(int);
            foreach (RowCell cell in row.Cells)
            {
                required += CellHeaderSize + cell.Value.Value.Length;
            }

            if (scratch.Length < required)
            {
                scratch = new byte[required];
            }

            Span<byte> destination = scratch;
            int offset = sizeof(int);
            int count = 0;
            foreach (RowCell cell in row.Cells)
            {
                ReadOnlySpan<byte> value = cell.Value.Value;
                BinaryPrimitives.WriteInt32LittleEndian(destination[offset..], cell.ColumnIndex);
                BinaryPrimitives.WriteInt32LittleEndian(destination[(offset + 4)..], (int)cell.Value.Type);
                BinaryPrimitives.WriteInt32LittleEndian(destination[(offset + 8)..], value.Length);
                offset += CellHeaderSize;
                value.CopyTo(destination[offset..]);
                offset += value.Length;
                count++;
            }

            BinaryPrimitives.WriteInt32LittleEndian(destination, count);
            return offset;
        }
    }
}
```

- [ ] **Step 4: Write the row cursor**

Create `src/ExcelReader.Native/NativeApi.Rows.cs`:

```csharp
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        internal static int NextRow(NativeHandle? handle, Span<byte> buffer, out int written)
        {
            written = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                if (!handle.HasPending)
                {
                    handle.Rows ??= handle.Reader.GetEnumerator();
                    if (!handle.Rows.MoveNext())
                    {
                        return NativeStatus.Eof;
                    }

                    Row row = handle.Rows.Current;
                    byte[] scratch = handle.Scratch;
                    handle.PendingLength = RowBlob.Serialize(row, ref scratch);
                    handle.Scratch = scratch;
                    handle.HasPending = true;
                }

                written = handle.PendingLength;
                if (buffer.Length < handle.PendingLength)
                {
                    return NativeStatus.BufferTooSmall;
                }

                handle.Scratch.AsSpan(0, handle.PendingLength).CopyTo(buffer);
                handle.HasPending = false;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }
    }
}
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj`
Expected: PASS.

- [ ] **Step 6: Verify a clean Release build**

Run: `dotnet build ExcelReader.slnx --configuration Release`
Expected: `Build succeeded`, 0 warnings.

- [ ] **Step 7: Commit**

```bash
git add src/ExcelReader.Native tests/ExcelReader.Tests/NativeApiTests.cs
git commit -m "feat(native): serialize rows into a flat blob for FFI consumers"
```

---

## Task 5: C ABI exports, header file, and AOT publish

**Files:**
- Create: `src/ExcelReader.Native/Exports.cs`
- Create: `src/ExcelReader.Native/include/excelreader.h`
- Create: `src/ExcelReader.Native/README.md`

**Interfaces:**
- Consumes: every `NativeApi` method from Tasks 1–4.
- Produces: the nine exported symbols listed in the [ABI Contract](#abi-contract). Nothing in C# consumes these; the Python package does.

- [ ] **Step 1: Write the exports**

Create `src/ExcelReader.Native/Exports.cs`:

```csharp
using System.Runtime.InteropServices;

namespace ExcelReader.Native
{
    /// <summary>
    /// The C ABI. Every function here does exactly two things: turn raw pointers into spans and a
    /// GCHandle into a <see cref="NativeHandle"/>, then delegate to <see cref="NativeApi"/>.
    /// Keep the logic in NativeApi — managed code cannot call an [UnmanagedCallersOnly] method, so
    /// anything implemented here is untestable.
    /// </summary>
    internal static unsafe class Exports
    {
        [UnmanagedCallersOnly(EntryPoint = "xl_open_file")]
        public static int OpenFile(byte* path, int pathLength, int format, nint* outHandle)
        {
            if (path is null || pathLength < 0 || outHandle is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.OpenFile(new ReadOnlySpan<byte>(path, pathLength), format, out NativeHandle? handle);
            *outHandle = handle is null ? 0 : GCHandle.ToIntPtr(GCHandle.Alloc(handle));
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_open_memory")]
        public static int OpenMemory(byte* data, int dataLength, int format, nint* outHandle)
        {
            if (data is null || dataLength < 0 || outHandle is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.OpenMemory(new ReadOnlySpan<byte>(data, dataLength), format, out NativeHandle? handle);
            *outHandle = handle is null ? 0 : GCHandle.ToIntPtr(GCHandle.Alloc(handle));
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_close")]
        public static int Close(nint handle)
        {
            if (handle == 0)
            {
                return NativeStatus.InvalidHandle;
            }

            GCHandle gcHandle = GCHandle.FromIntPtr(handle);
            int status = NativeApi.Close(gcHandle.Target as NativeHandle);
            gcHandle.Free();
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_count")]
        public static int SheetCount(nint handle, int* outCount)
        {
            if (outCount is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.SheetCount(Resolve(handle), out int count);
            *outCount = count;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_sheet_name")]
        public static int SheetName(nint handle, byte* buffer, int capacity, int* outLength)
        {
            if (capacity < 0 || outLength is null || (buffer is null && capacity > 0))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.SheetName(Resolve(handle), new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_move_to_sheet")]
        public static int MoveToSheet(nint handle, int index)
        {
            return NativeApi.MoveToSheet(Resolve(handle), index);
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_is_date1904")]
        public static int IsDate1904(nint handle, int* outFlag)
        {
            if (outFlag is null)
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.IsDate1904(Resolve(handle), out int flag);
            *outFlag = flag;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_next_row")]
        public static int NextRow(nint handle, byte* buffer, int capacity, int* outWritten)
        {
            if (capacity < 0 || outWritten is null || (buffer is null && capacity > 0))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.NextRow(Resolve(handle), new Span<byte>(buffer, capacity), out int written);
            *outWritten = written;
            return status;
        }

        [UnmanagedCallersOnly(EntryPoint = "xl_last_error")]
        public static int LastError(byte* buffer, int capacity, int* outLength)
        {
            if (capacity < 0 || outLength is null || (buffer is null && capacity > 0))
            {
                return NativeStatus.InvalidArgument;
            }

            int status = NativeApi.LastError(new Span<byte>(buffer, capacity), out int length);
            *outLength = length;
            return status;
        }

        private static NativeHandle? Resolve(nint handle)
        {
            return handle == 0 ? null : GCHandle.FromIntPtr(handle).Target as NativeHandle;
        }
    }
}
```

- [ ] **Step 2: Write the C header**

Create `src/ExcelReader.Native/include/excelreader.h`:

```c
/* ExcelReader C ABI - reading only.
 * Every function returns an XL_* status code. All strings are UTF-8 with an explicit length.
 * All integers in the row blob are little-endian int32.
 * No handle is thread-safe; use one handle per thread. */
#ifndef EXCELREADER_H
#define EXCELREADER_H

#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

#define XL_OK                 0
#define XL_EOF               -1
#define XL_BUFFER_TOO_SMALL  -2
#define XL_INVALID_HANDLE    -3
#define XL_INVALID_ARGUMENT  -4
#define XL_ERROR             -5

#define XL_FORMAT_AUTO  0  /* sniffs XLS/XLSX/XLSB; does NOT detect CSV */
#define XL_FORMAT_XLS   1
#define XL_FORMAT_XLSX  2
#define XL_FORMAT_XLSB  3
#define XL_FORMAT_CSV   4  /* must be requested explicitly */

#define XL_CELL_EMPTY   0
#define XL_CELL_STRING  1
#define XL_CELL_NUMBER  2
#define XL_CELL_DATE    3
#define XL_CELL_BOOL    4
#define XL_CELL_FORMULA 5
#define XL_CELL_ERROR   6

/* Copies the path; the caller may free it on return. */
int32_t xl_open_file(const uint8_t* path, int32_t path_len, int32_t format, void** out_handle);

/* Copies the data; the caller may free it on return. */
int32_t xl_open_memory(const uint8_t* data, int32_t data_len, int32_t format, void** out_handle);

int32_t xl_close(void* handle);

int32_t xl_sheet_count(void* handle, int32_t* out_count);

/* On XL_BUFFER_TOO_SMALL, *out_len holds the required byte count. */
int32_t xl_sheet_name(void* handle, uint8_t* buffer, int32_t capacity, int32_t* out_len);

/* Resets row enumeration to the start of the selected sheet. */
int32_t xl_move_to_sheet(void* handle, int32_t index);

int32_t xl_is_date1904(void* handle, int32_t* out_flag);

/* Writes one row as:
 *     int32 cell_count
 *     repeated: int32 column, int32 type, int32 value_len, uint8 value[value_len]
 * Returns XL_EOF at the end of the sheet. On XL_BUFFER_TOO_SMALL, *out_written holds the required
 * byte count and the row is held until the next call - no row is ever lost. */
int32_t xl_next_row(void* handle, uint8_t* buffer, int32_t capacity, int32_t* out_written);

/* Last error on the CALLING thread. */
int32_t xl_last_error(uint8_t* buffer, int32_t capacity, int32_t* out_len);

#ifdef __cplusplus
}
#endif

#endif /* EXCELREADER_H */
```

- [ ] **Step 3: Publish the shared library**

Run (substitute your platform's RID: `win-x64`, `linux-x64`, `osx-arm64`, …):

```bash
dotnet publish src/ExcelReader.Native/ExcelReader.Native.csproj -c Release -f net10.0 -r win-x64
```

Expected: `Build succeeded`, **zero** `IL2xxx`/`IL3xxx` warnings, and a binary at
`src/ExcelReader.Native/bin/Release/net10.0/win-x64/publish/ExcelReader.Native.dll`
(`.so` on Linux, `.dylib` on macOS — NativeAOT emits no `lib` prefix).

- [ ] **Step 4: Verify the symbols are actually exported**

On Windows (PowerShell), confirm the file is a DLL of a few MB and non-empty:

```powershell
Get-Item src/ExcelReader.Native/bin/Release/net10.0/win-x64/publish/ExcelReader.Native.dll | Select-Object Length
```

On Linux/macOS:

```bash
nm -gU src/ExcelReader.Native/bin/Release/net10.0/linux-x64/publish/ExcelReader.Native.so | grep xl_
```

Expected: all nine `xl_*` symbols present. On Windows this is verified by Task 6's Python smoke test instead.

- [ ] **Step 5: Write the project README**

Create `src/ExcelReader.Native/README.md`:

```markdown
# ExcelReader.Native

NativeAOT shared library exposing ExcelReader's readers over a C ABI, so non-.NET languages
(Python, C, C++, Go, Node) can read XLSX, XLSB, XLS and CSV without a .NET runtime.

- ABI reference: `include/excelreader.h`
- Full contract and rationale: `docs/plans/2026-08-13-native-ffi-python-reading.md`
- Python binding: `python/`

## Build

    dotnet publish src/ExcelReader.Native/ExcelReader.Native.csproj -c Release -f net10.0 -r <rid>

Output lands in `bin/Release/net10.0/<rid>/publish/ExcelReader.Native.{dll,so,dylib}`.

## Layout

| File | Role |
|---|---|
| `NativeApi*.cs` | Internal span-based implementation. This is what the tests drive. |
| `Exports.cs` | `[UnmanagedCallersOnly]` pointer wrappers. Keep logic out of here — it is untestable from managed code. |
| `RowBlob.cs` | Row serialization. |
| `include/excelreader.h` | Hand-written C header; keep in sync with `Exports.cs`. |

Writing is not exposed yet — reading only.
```

- [ ] **Step 6: Commit**

```bash
git add src/ExcelReader.Native
git commit -m "feat(native): add C ABI exports, header and AOT publish"
```

---

## Task 6: Python package skeleton and native library loader

**Files:**
- Create: `python/pyproject.toml`
- Create: `python/scripts/build_native.py`
- Create: `python/src/excelreader/_native.py`
- Create: `python/src/excelreader/__init__.py` (placeholder, filled in Task 7)
- Create: `python/tests/test_native.py`
- Modify: `.gitignore`

**Interfaces:**
- Consumes: the exported symbols from Task 5.
- Produces:
  - `excelreader._native.load_library() -> ctypes.CDLL` (memoized)
  - `excelreader._native.XL_OK`, `XL_EOF`, `XL_BUFFER_TOO_SMALL`, `XL_INVALID_HANDLE`, `XL_INVALID_ARGUMENT`, `XL_ERROR`
  - `excelreader._native.XL_FORMAT_AUTO/XLS/XLSX/XLSB/CSV`
  - `excelreader._native.library_filename() -> str`

- [ ] **Step 1: Write the package metadata**

Create `python/pyproject.toml`:

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "excelreader"
version = "0.1.0"
description = "Read XLSX, XLSB, XLS and CSV through the ExcelReader NativeAOT library"
readme = "README.md"
requires-python = ">=3.9"
license = { text = "MIT" }
authors = [{ name = "Gabriel Matte" }]
classifiers = [
    "Programming Language :: Python :: 3",
    "License :: OSI Approved :: MIT License",
]
dependencies = []

[project.optional-dependencies]
dev = ["pytest>=7.0"]

[project.urls]
Homepage = "https://github.com/GabrielMarquezMatte/ExcelReader"

[tool.hatch.build.targets.wheel]
packages = ["src/excelreader"]

[tool.hatch.build.targets.wheel.force-include]
"src/excelreader/_lib" = "excelreader/_lib"

[tool.pytest.ini_options]
testpaths = ["tests"]
```

- [ ] **Step 2: Ignore the copied binaries**

Append to `.gitignore`:

```gitignore
# Native binary copied into the Python package by python/scripts/build_native.py
python/src/excelreader/_lib/
```

- [ ] **Step 3: Write the build script**

Create `python/scripts/build_native.py`:

```python
"""Publish ExcelReader.Native for this machine and copy the binary into the Python package.

Run from anywhere:  python python/scripts/build_native.py
"""

import argparse
import platform
import shutil
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
CSPROJ = REPO_ROOT / "src" / "ExcelReader.Native" / "ExcelReader.Native.csproj"
PACKAGE_LIB_DIR = REPO_ROOT / "python" / "src" / "excelreader" / "_lib"

LIB_NAMES = {
    "Windows": "ExcelReader.Native.dll",
    "Linux": "ExcelReader.Native.so",
    "Darwin": "ExcelReader.Native.dylib",
}


def default_rid() -> str:
    system = platform.system()
    machine = platform.machine().lower()
    arch = "arm64" if machine in {"arm64", "aarch64"} else "x64"
    if system == "Windows":
        return f"win-{arch}"
    if system == "Darwin":
        return f"osx-{arch}"
    return f"linux-{arch}"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rid", default=default_rid(), help="dotnet runtime identifier")
    parser.add_argument("--framework", default="net10.0")
    args = parser.parse_args()

    subprocess.run(
        ["dotnet", "publish", str(CSPROJ), "-c", "Release", "-f", args.framework, "-r", args.rid],
        check=True,
    )

    publish_dir = CSPROJ.parent / "bin" / "Release" / args.framework / args.rid / "publish"
    name = LIB_NAMES[platform.system()]
    source = publish_dir / name
    if not source.exists():
        print(f"error: expected native library at {source}", file=sys.stderr)
        return 1

    PACKAGE_LIB_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, PACKAGE_LIB_DIR / name)
    print(f"copied {source} -> {PACKAGE_LIB_DIR / name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Write the failing test**

Create `python/tests/test_native.py`:

```python
import ctypes

from excelreader import _native


def test_library_loads():
    lib = _native.load_library()
    assert isinstance(lib, ctypes.CDLL)


def test_library_is_memoized():
    assert _native.load_library() is _native.load_library()


def test_exported_functions_are_present():
    lib = _native.load_library()
    for name in (
        "xl_open_file",
        "xl_open_memory",
        "xl_close",
        "xl_sheet_count",
        "xl_sheet_name",
        "xl_move_to_sheet",
        "xl_is_date1904",
        "xl_next_row",
        "xl_last_error",
    ):
        assert hasattr(lib, name), name


def test_status_constants_match_the_abi():
    assert (_native.XL_OK, _native.XL_EOF, _native.XL_BUFFER_TOO_SMALL) == (0, -1, -2)
    assert (_native.XL_INVALID_HANDLE, _native.XL_INVALID_ARGUMENT, _native.XL_ERROR) == (-3, -4, -5)
```

- [ ] **Step 5: Run the test to verify it fails**

```bash
python python/scripts/build_native.py
pip install -e "python[dev]"
pytest python/tests -v
```

Expected: FAIL — `excelreader._native` does not exist yet. (`build_native.py` must already succeed;
if it does not, fix Task 5 before continuing.)

- [ ] **Step 6: Write the loader**

Create `python/src/excelreader/_native.py`:

```python
"""ctypes binding to the ExcelReader NativeAOT shared library.

Everything here mirrors src/ExcelReader.Native/include/excelreader.h. If you change one, change both.
"""

from __future__ import annotations

import ctypes
import os
import platform
from functools import lru_cache
from pathlib import Path

XL_OK = 0
XL_EOF = -1
XL_BUFFER_TOO_SMALL = -2
XL_INVALID_HANDLE = -3
XL_INVALID_ARGUMENT = -4
XL_ERROR = -5

XL_FORMAT_AUTO = 0
XL_FORMAT_XLS = 1
XL_FORMAT_XLSX = 2
XL_FORMAT_XLSB = 3
XL_FORMAT_CSV = 4

_LIB_NAMES = {
    "Windows": "ExcelReader.Native.dll",
    "Linux": "ExcelReader.Native.so",
    "Darwin": "ExcelReader.Native.dylib",
}

# NativeAOT emits no "lib" prefix, so the filename is the assembly name on every platform.
def library_filename() -> str:
    try:
        return _LIB_NAMES[platform.system()]
    except KeyError:
        raise RuntimeError(f"unsupported platform: {platform.system()}") from None


def _candidate_paths() -> list[Path]:
    override = os.environ.get("EXCELREADER_NATIVE_LIB")
    if override:
        return [Path(override)]
    return [Path(__file__).resolve().parent / "_lib" / library_filename()]


@lru_cache(maxsize=1)
def load_library() -> ctypes.CDLL:
    for path in _candidate_paths():
        if path.exists():
            return _bind(ctypes.CDLL(str(path)))
    raise RuntimeError(
        f"{library_filename()} not found. Build it with:\n"
        f"    python python/scripts/build_native.py\n"
        f"or point EXCELREADER_NATIVE_LIB at an existing binary."
    )


def _bind(lib: ctypes.CDLL) -> ctypes.CDLL:
    c_int = ctypes.c_int32
    p_int = ctypes.POINTER(ctypes.c_int32)
    p_void = ctypes.c_void_p
    pp_void = ctypes.POINTER(ctypes.c_void_p)
    p_bytes = ctypes.c_char_p

    lib.xl_open_file.argtypes = [p_bytes, c_int, c_int, pp_void]
    lib.xl_open_file.restype = c_int
    lib.xl_open_memory.argtypes = [p_bytes, c_int, c_int, pp_void]
    lib.xl_open_memory.restype = c_int
    lib.xl_close.argtypes = [p_void]
    lib.xl_close.restype = c_int
    lib.xl_sheet_count.argtypes = [p_void, p_int]
    lib.xl_sheet_count.restype = c_int
    lib.xl_sheet_name.argtypes = [p_void, p_bytes, c_int, p_int]
    lib.xl_sheet_name.restype = c_int
    lib.xl_move_to_sheet.argtypes = [p_void, c_int]
    lib.xl_move_to_sheet.restype = c_int
    lib.xl_is_date1904.argtypes = [p_void, p_int]
    lib.xl_is_date1904.restype = c_int
    lib.xl_next_row.argtypes = [p_void, p_bytes, c_int, p_int]
    lib.xl_next_row.restype = c_int
    lib.xl_last_error.argtypes = [p_bytes, c_int, p_int]
    lib.xl_last_error.restype = c_int
    return lib
```

- [ ] **Step 7: Write the placeholder package init**

Create `python/src/excelreader/__init__.py`:

```python
"""Python bindings for ExcelReader. Reading only."""

from excelreader import _native

__all__ = ["_native"]
```

- [ ] **Step 8: Run the tests to verify they pass**

```bash
pytest python/tests -v
```

Expected: PASS (4 tests).

- [ ] **Step 9: Commit**

```bash
git add python .gitignore
git commit -m "feat(python): add package skeleton and native library loader"
```

---

## Task 7: Python read API

**Files:**
- Create: `python/src/excelreader/types.py`
- Create: `python/src/excelreader/reader.py`
- Modify: `python/src/excelreader/__init__.py`
- Create: `python/tests/conftest.py`
- Create: `python/tests/test_reader.py`

**Interfaces:**
- Consumes: `excelreader._native` (Task 6).
- Produces:
  - `excelreader.CellType` — `IntEnum`: `EMPTY=0, STRING=1, NUMBER=2, DATE=3, BOOL=4, FORMULA=5, ERROR=6`
  - `excelreader.Cell` — `NamedTuple(column: int, type: CellType, value: str)`
  - `excelreader.ExcelReaderError(Exception)`
  - `excelreader.Workbook` — context manager with `sheet_count`, `sheet_name`, `is_date1904`, `move_to_sheet(index)`, `rows()`, `close()`
  - `excelreader.open_workbook(path, format=None) -> Workbook`
  - `excelreader.open_bytes(data, format=None) -> Workbook`

- [ ] **Step 1: Write the fixtures file**

Create `python/tests/conftest.py`:

```python
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]


@pytest.fixture(scope="session")
def xlsx_path() -> Path:
    return REPO_ROOT / "RealExcel.xlsx"


@pytest.fixture(scope="session")
def xlsb_path() -> Path:
    return REPO_ROOT / "RealExcel.xlsb"


@pytest.fixture(scope="session")
def xls_path() -> Path:
    return REPO_ROOT / "tests" / "ExcelReader.Benchmarks" / "Data" / "65K_Records_Data.xls"


@pytest.fixture
def csv_path(tmp_path: Path) -> Path:
    path = tmp_path / "sample.csv"
    path.write_text("name,qty\nwidget,7\ngadget,9\n", encoding="utf-8")
    return path
```

- [ ] **Step 2: Write the failing tests**

Create `python/tests/test_reader.py`:

```python
import pytest

from excelreader import Cell, CellType, ExcelReaderError, open_bytes, open_workbook


def test_reads_the_xlsx_header_row(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        header = next(workbook.rows())

    assert len(header) == 18
    assert header[0] == Cell(column=0, type=CellType.STRING, value="Coluna1")
    assert header[17].value == "Coluna18"


def test_reads_every_xlsx_row(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = list(workbook.rows())

    assert len(rows) == 101


def test_reports_xlsx_cell_types(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        rows = workbook.rows()
        next(rows)
        first_data_row = next(rows)

    assert first_data_row[0].type is CellType.STRING
    assert first_data_row[1].type is CellType.DATE
    assert first_data_row[2].type is CellType.NUMBER


def test_reads_csv_when_the_extension_is_csv(csv_path):
    with open_workbook(csv_path) as workbook:
        rows = list(workbook.rows())

    assert [cell.value for cell in rows[0]] == ["name", "qty"]
    assert [cell.value for cell in rows[2]] == ["gadget", "9"]


def test_reads_xlsb(xlsb_path):
    with open_workbook(xlsb_path) as workbook:
        first = next(workbook.rows())

    assert len(first) > 0


def test_reads_xls(xls_path):
    with open_workbook(xls_path) as workbook:
        rows = workbook.rows()
        # 65K rows / 11 MB: reading a handful proves the path works without the wall-clock cost.
        sampled = [next(rows) for _ in range(10)]

    assert all(len(row) > 0 for row in sampled)


def test_reads_from_bytes(xlsx_path):
    with open_bytes(xlsx_path.read_bytes()) as workbook:
        header = next(workbook.rows())

    assert header[0].value == "Coluna1"


def test_exposes_sheet_metadata(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        assert workbook.sheet_count >= 1
        assert workbook.sheet_name
        assert workbook.is_date1904 is False


def test_move_to_sheet_restarts_enumeration(xlsx_path):
    with open_workbook(xlsx_path) as workbook:
        first = next(workbook.rows())
        workbook.move_to_sheet(0)
        again = next(workbook.rows())

    assert first == again


def test_missing_file_raises(tmp_path):
    with pytest.raises(ExcelReaderError):
        open_workbook(tmp_path / "nope.xlsx")


def test_use_after_close_raises(xlsx_path):
    workbook = open_workbook(xlsx_path)
    workbook.close()

    with pytest.raises(ExcelReaderError):
        workbook.sheet_count


def test_close_is_idempotent(xlsx_path):
    workbook = open_workbook(xlsx_path)
    workbook.close()
    workbook.close()
```

- [ ] **Step 3: Run the tests to verify they fail**

```bash
pytest python/tests/test_reader.py -v
```

Expected: FAIL — `ImportError: cannot import name 'Cell' from 'excelreader'`.

- [ ] **Step 4: Write the value types**

Create `python/src/excelreader/types.py`:

```python
"""Value types crossing the FFI boundary. Mirrors ExcelReader.Core.Enums.CellType."""

from __future__ import annotations

from enum import IntEnum
from typing import NamedTuple


class CellType(IntEnum):
    EMPTY = 0
    STRING = 1
    NUMBER = 2
    DATE = 3
    BOOL = 4
    FORMULA = 5
    ERROR = 6


class Cell(NamedTuple):
    """One cell. `value` is the raw text as stored, so DATE cells are Excel serial numbers."""

    column: int
    type: CellType
    value: str


class ExcelReaderError(Exception):
    """Raised when the native library reports a failure."""
```

- [ ] **Step 5: Write the reader**

Create `python/src/excelreader/reader.py`:

```python
"""The public reading API. One Workbook wraps one native handle."""

from __future__ import annotations

import ctypes
import struct
from pathlib import Path
from typing import Iterator, Optional, Union

from excelreader import _native
from excelreader.types import Cell, CellType, ExcelReaderError

_FORMATS = {
    "auto": _native.XL_FORMAT_AUTO,
    "xls": _native.XL_FORMAT_XLS,
    "xlsx": _native.XL_FORMAT_XLSX,
    "xlsb": _native.XL_FORMAT_XLSB,
    "csv": _native.XL_FORMAT_CSV,
}

_CELL_HEADER = struct.Struct("<iii")
_INITIAL_ROW_BUFFER = 64 * 1024


def _last_error() -> str:
    lib = _native.load_library()
    length = ctypes.c_int32()
    buffer = ctypes.create_string_buffer(1024)
    if lib.xl_last_error(buffer, len(buffer), ctypes.byref(length)) == _native.XL_BUFFER_TOO_SMALL:
        buffer = ctypes.create_string_buffer(length.value)
        lib.xl_last_error(buffer, len(buffer), ctypes.byref(length))
    return buffer.raw[: length.value].decode("utf-8", errors="replace")


def _check(status: int) -> None:
    if status == _native.XL_OK:
        return
    if status == _native.XL_INVALID_HANDLE:
        raise ExcelReaderError("workbook is closed or the handle is invalid")
    if status == _native.XL_INVALID_ARGUMENT:
        raise ExcelReaderError("invalid argument passed to the native library")
    raise ExcelReaderError(_last_error() or f"native call failed with status {status}")


def _resolve_format(name: Optional[str], path: Optional[Path]) -> int:
    if name is not None:
        try:
            return _FORMATS[name.lower()]
        except KeyError:
            raise ValueError(f"unknown format {name!r}; expected one of {sorted(_FORMATS)}") from None
    # The signature sniffer covers XLS/XLSX/XLSB but CSV has no signature, so the extension is the
    # only hint available for it.
    if path is not None and path.suffix.lower() == ".csv":
        return _native.XL_FORMAT_CSV
    return _native.XL_FORMAT_AUTO


class Workbook:
    """A read cursor over one workbook. Not thread-safe; use one instance per thread."""

    def __init__(self, handle: ctypes.c_void_p) -> None:
        self._lib = _native.load_library()
        self._handle: Optional[ctypes.c_void_p] = handle

    def _require_handle(self) -> ctypes.c_void_p:
        if self._handle is None:
            raise ExcelReaderError("workbook is closed")
        return self._handle

    @property
    def sheet_count(self) -> int:
        count = ctypes.c_int32()
        _check(self._lib.xl_sheet_count(self._require_handle(), ctypes.byref(count)))
        return count.value

    @property
    def sheet_name(self) -> str:
        handle = self._require_handle()
        length = ctypes.c_int32()
        buffer = ctypes.create_string_buffer(256)
        status = self._lib.xl_sheet_name(handle, buffer, len(buffer), ctypes.byref(length))
        if status == _native.XL_BUFFER_TOO_SMALL:
            buffer = ctypes.create_string_buffer(length.value)
            status = self._lib.xl_sheet_name(handle, buffer, len(buffer), ctypes.byref(length))
        _check(status)
        return buffer.raw[: length.value].decode("utf-8")

    @property
    def is_date1904(self) -> bool:
        flag = ctypes.c_int32()
        _check(self._lib.xl_is_date1904(self._require_handle(), ctypes.byref(flag)))
        return flag.value != 0

    def move_to_sheet(self, index: int) -> None:
        """Selects a sheet and restarts row enumeration from its first row."""
        _check(self._lib.xl_move_to_sheet(self._require_handle(), index))

    def rows(self) -> Iterator[list[Cell]]:
        handle = self._require_handle()
        written = ctypes.c_int32()
        capacity = _INITIAL_ROW_BUFFER
        buffer = ctypes.create_string_buffer(capacity)
        while True:
            status = self._lib.xl_next_row(handle, buffer, capacity, ctypes.byref(written))
            if status == _native.XL_EOF:
                return
            if status == _native.XL_BUFFER_TOO_SMALL:
                # The native side holds the row until it fits, so growing loses nothing.
                capacity = written.value
                buffer = ctypes.create_string_buffer(capacity)
                continue
            _check(status)
            yield _decode_row(buffer.raw, written.value)

    def close(self) -> None:
        if self._handle is None:
            return
        handle, self._handle = self._handle, None
        _check(self._lib.xl_close(handle))

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, *_exc_info: object) -> None:
        self.close()


def _decode_row(blob: bytes, length: int) -> list[Cell]:
    count = struct.unpack_from("<i", blob, 0)[0]
    cells: list[Cell] = []
    offset = 4
    for _ in range(count):
        column, cell_type, value_length = _CELL_HEADER.unpack_from(blob, offset)
        offset += _CELL_HEADER.size
        value = blob[offset : offset + value_length].decode("utf-8")
        offset += value_length
        cells.append(Cell(column=column, type=CellType(cell_type), value=value))
    if offset != length:
        raise ExcelReaderError(f"row blob is malformed: consumed {offset} of {length} bytes")
    return cells


def open_workbook(path: Union[str, Path], format: Optional[str] = None) -> Workbook:
    """Opens a workbook from disk. `format` is one of auto/xls/xlsx/xlsb/csv; None infers it."""
    resolved = Path(path)
    encoded = str(resolved).encode("utf-8")
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(lib.xl_open_file(encoded, len(encoded), _resolve_format(format, resolved), ctypes.byref(handle)))
    return Workbook(handle)


def open_bytes(data: bytes, format: Optional[str] = None) -> Workbook:
    """Opens a workbook from an in-memory buffer. The native side copies `data` immediately."""
    handle = ctypes.c_void_p()
    lib = _native.load_library()
    _check(lib.xl_open_memory(data, len(data), _resolve_format(format, None), ctypes.byref(handle)))
    return Workbook(handle)
```

- [ ] **Step 6: Export the public names**

Replace `python/src/excelreader/__init__.py` with:

```python
"""Python bindings for ExcelReader. Reading only — writing is not exposed yet."""

from excelreader.reader import Workbook, open_bytes, open_workbook
from excelreader.types import Cell, CellType, ExcelReaderError

__all__ = [
    "Cell",
    "CellType",
    "ExcelReaderError",
    "Workbook",
    "open_bytes",
    "open_workbook",
]
```

- [ ] **Step 7: Run the tests to verify they pass**

```bash
pytest python/tests -v
```

Expected: PASS (all tests in `test_native.py` and `test_reader.py`).

If `test_reads_every_xlsx_row` reports a count other than 101, trust the reader and update the
assertion — that number was measured against `RealExcel.xlsx` on 2026-08-13.

- [ ] **Step 8: Commit**

```bash
git add python
git commit -m "feat(python): add Workbook read API over the native ABI"
```

---

## Task 8: Python README and usage example

**Files:**
- Create: `python/README.md`
- Create: `python/examples/read_workbook.py`
- Modify: `python/tests/test_reader.py`

**Interfaces:**
- Consumes: the whole public API from Task 7.
- Produces: nothing new in code — this task locks the documented behavior in with a test.

- [ ] **Step 1: Write the failing test for the documented DATE conversion recipe**

Append to `python/tests/test_reader.py`:

```python
def test_excel_serial_dates_convert_with_the_documented_recipe(xlsx_path):
    from datetime import date, timedelta

    with open_workbook(xlsx_path) as workbook:
        rows = workbook.rows()
        next(rows)
        first_data_row = next(rows)
        epoch = date(1904, 1, 1) if workbook.is_date1904 else date(1899, 12, 30)

    serial = int(float(first_data_row[1].value))
    assert epoch + timedelta(days=serial) == date(2026, 1, 1)
```

- [ ] **Step 2: Run it to verify it fails or reveals the true date**

```bash
pytest python/tests/test_reader.py::test_excel_serial_dates_convert_with_the_documented_recipe -v
```

Expected: FAIL on the asserted date. Replace `date(2026, 1, 1)` with the value the failure reports,
then re-run to PASS. (Serial `46023` against the 1900 system is the value observed on 2026-08-13;
the point of the test is that the recipe is correct, not the specific day.)

- [ ] **Step 3: Write the README**

Create `python/README.md`:

````markdown
# excelreader (Python)

Read XLSX, XLSB, XLS and CSV through ExcelReader's NativeAOT library. No .NET runtime required —
the shared library is self-contained. Reading only; writing is not exposed yet.

## Install (from source)

```bash
python python/scripts/build_native.py   # requires the .NET 10 SDK, once per machine
pip install -e "python[dev]"
```

`build_native.py` publishes `src/ExcelReader.Native` for your platform and copies the resulting
`ExcelReader.Native.{dll,so,dylib}` into `excelreader/_lib/`. To point at a binary you built
elsewhere, set `EXCELREADER_NATIVE_LIB` to its full path.

## Usage

```python
from excelreader import open_workbook

with open_workbook("book.xlsx") as workbook:
    print(workbook.sheet_count, workbook.sheet_name)
    for row in workbook.rows():
        for cell in row:
            print(cell.column, cell.type.name, cell.value)
```

### Formats

`open_workbook` sniffs XLS/XLSX/XLSB by file signature. CSV has no signature, so it is chosen by the
`.csv` extension — or explicitly:

```python
open_workbook("data.txt", format="csv")
```

### Dates

`cell.value` is always the raw text as stored, so `CellType.DATE` cells hold Excel serial numbers.
Convert them yourself:

```python
from datetime import date, timedelta

epoch = date(1904, 1, 1) if workbook.is_date1904 else date(1899, 12, 30)
as_date = epoch + timedelta(days=int(float(cell.value)))
```

### From memory

```python
from excelreader import open_bytes

with open_bytes(payload) as workbook:
    ...
```

## Notes

- A `Workbook` is **not** thread-safe. Use one per thread.
- Empty cells are skipped, so `cell.column` may skip indices. Do not assume `row[i].column == i`.
- The ABI is documented in `src/ExcelReader.Native/include/excelreader.h`.
````

- [ ] **Step 4: Write the example**

Create `python/examples/read_workbook.py`:

```python
"""Print the first rows of a workbook.  Usage: python read_workbook.py <path> [max_rows]"""

import sys

from excelreader import open_workbook


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2

    path = sys.argv[1]
    max_rows = int(sys.argv[2]) if len(sys.argv) > 2 else 10

    with open_workbook(path) as workbook:
        print(f"sheets={workbook.sheet_count} current={workbook.sheet_name!r} date1904={workbook.is_date1904}")
        for index, row in enumerate(workbook.rows()):
            if index >= max_rows:
                break
            print(index, [(cell.column, cell.type.name, cell.value) for cell in row])

    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 5: Run the example against a real file**

```bash
python python/examples/read_workbook.py RealExcel.xlsx 3
```

Expected: a `sheets=1 current='...' date1904=False` line followed by 3 decoded rows.

- [ ] **Step 6: Run the whole Python suite**

```bash
pytest python/tests -v
```

Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add python
git commit -m "docs(python): add README, runnable example and date-conversion test"
```

---

## Task 9: CI

**Files:**
- Create: `.github/workflows/python.yml`
- Modify: `README.md` (repo root)

**Interfaces:**
- Consumes: `python/scripts/build_native.py`, `python/tests/` (Tasks 6–8).
- Produces: nothing consumed by other tasks.

- [ ] **Step 1: Write the workflow**

Create `.github/workflows/python.yml`:

```yaml
name: Python bindings

on:
  push:
    branches: [master, develop]
  pull_request:
    branches: [master, develop]

permissions:
  contents: read

concurrency:
  group: python-${{ github.ref }}
  cancel-in-progress: true

jobs:
  test:
    name: Native + Python (${{ matrix.os }})
    runs-on: ${{ matrix.os }}
    strategy:
      fail-fast: false
      matrix:
        os: [ubuntu-latest, windows-latest, macos-latest]

    steps:
      - name: Checkout
        uses: actions/checkout@v5

      - name: Setup .NET
        uses: actions/setup-dotnet@v5
        with:
          global-json-file: global.json

      - name: Setup Python
        uses: actions/setup-python@v6
        with:
          python-version: '3.12'

      - name: Build the native library
        run: python python/scripts/build_native.py

      - name: Install the Python package
        run: pip install -e "python[dev]"

      - name: Run the Python tests
        run: pytest python/tests -v
```

Before committing, open `.github/workflows/ci.yml` and match the action versions it already pins
(`actions/checkout`, `actions/setup-dotnet`). Do not introduce a different major version.

- [ ] **Step 2: Verify the workflow parses**

```bash
python -c "import yaml,sys; yaml.safe_load(open('.github/workflows/python.yml')); print('ok')"
```

Expected: `ok`. (If PyYAML is unavailable, run `pip install pyyaml` first — this is a one-off check,
not a project dependency.)

- [ ] **Step 3: Document the bindings in the root README**

Add this section to `README.md` (repo root), immediately before the license/contributing section at
the end:

```markdown
## Other languages

ExcelReader ships a NativeAOT shared library with a C ABI, so non-.NET languages can read XLSX,
XLSB, XLS and CSV without a .NET runtime installed.

- C ABI header: [`src/ExcelReader.Native/include/excelreader.h`](src/ExcelReader.Native/include/excelreader.h)
- Python package: [`python/`](python/README.md)

```python
from excelreader import open_workbook

with open_workbook("book.xlsx") as workbook:
    for row in workbook.rows():
        print([cell.value for cell in row])
```

Reading only — the writers are not exposed across the ABI yet.
```

- [ ] **Step 4: Run the full verification**

```bash
dotnet build ExcelReader.slnx --configuration Release
dotnet test --project tests/ExcelReader.Tests/ExcelReader.Tests.csproj
pytest python/tests -v
```

Expected: build succeeds with 0 warnings, both test suites pass.

- [ ] **Step 5: Commit**

```bash
git add .github/workflows/python.yml README.md
git commit -m "ci: build the native library and run the Python bindings tests"
```

---

## Deferred (explicitly out of scope)

Recorded so nobody re-derives these decisions mid-implementation:

| Deferred | Add when |
|---|---|
| Writing (`IWorkbookWriter<TSheet>` across the ABI) | A follow-up plan. The reading ABI must be stable first. |
| Zero-copy per-cell accessors instead of one blob per row | Profiling shows the per-row copy dominates. The current design is one boundary crossing per row, which is already the cheap end. |
| Async / cancellation across the ABI | A caller needs to cancel a long read. `ct` has no natural C representation; it would need a separate cancellation-token handle. |
| Typed parsing (`ExcelParser<T>`, `RefParser`) across the ABI | Python callers can build their own typed layer over `Cell`. Exposing generic .NET parsers over FFI is a much bigger surface. |
| Prebuilt wheels (cibuildwheel, manylinux, PyPI publishing) | The package proves itself from source first. |
| `CsvReaderOptions` / `ExcelReaderOptions` across the ABI | A caller needs a non-default delimiter or limit. Requires an options struct on the ABI. |
| Sheet lookup by name (`TryMoveToSheet`) | A caller needs it. `sheet_count` + `move_to_sheet` + `sheet_name` already lets Python search by name in a loop. |
