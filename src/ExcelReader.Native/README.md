# ExcelReader.Native

NativeAOT shared library exposing ExcelReader's readers over a C ABI, so non-.NET languages
(Python, C, C++, Go, Node) can read XLSX, XLSB, XLS and CSV without a .NET runtime.

- ABI reference: `include/excelreader.h`
- Full contract and rationale: `docs/NATIVE_BINDINGS_PLAN.md`, `docs/NATIVE_HARDENING_PLAN.md`
- Python binding: `python/`

## Build

    dotnet publish src/ExcelReader.Native/ExcelReader.Native.csproj -c Release -f net10.0 -r <rid>

Output lands in `bin/Release/net10.0/<rid>/publish/ExcelReader.Native.{dll,so,dylib}`.

## Layout

| File | Role |
|---|---|
| `NativeApi*.cs` | Internal span-based implementation. This is what the tests drive. |
| `Exports.cs` | `[UnmanagedCallersOnly]` pointer wrappers. Keep logic out of here — it is untestable from managed code. |
| `NativeHandleTable.cs` | Maps the opaque handle ids callers see onto `NativeHandle` instances. Ids are never reissued after `xl_close`, so a stale handle stays invalid permanently. |
| `RowBlob.cs` | Row serialization. |
| `include/excelreader.h` | Hand-written C header; keep in sync with `Exports.cs`. |

Writing is not exposed yet — reading only.

## Consuming from C

Include `include/excelreader.h` (add `include/excelreader_arrow.h` too if you want the Arrow C Data
Interface export). Every caller should check `xl_abi_version()` against `XL_ABI_VERSION` before doing
anything else and refuse to proceed on a mismatch:

```c
#include "excelreader.h"

if (xl_abi_version() != XL_ABI_VERSION) {
    /* rebuild against a matching header, or refuse to proceed */
}
```

**Linking.** NativeAOT's publish output ships no `ExcelReader.Native.lib` import library on Windows,
so a normal link step against the DLL does not work with MSVC out of the box. The verified, portable
approach — used by `tests/ExcelReader.NativeSmoke/`, the complete worked example — is to load the
shared library dynamically instead of linking it at build time:

- Windows: `LoadLibraryA` + `GetProcAddress`
- Linux/macOS: `dlopen` + `dlsym` (link `libdl` on Linux; already part of libc on macOS)

If your toolchain can produce an import library for the DLL (e.g. via a `.def` file and `lib.exe`), a
normal link also works — this has not been set up or verified here.

**Verifying it works:** `tests/ExcelReader.NativeSmoke/` is the reference consumer — it exercises
every export against `RealExcel.xlsb` and pins the ABI structs' layout with `_STATIC_ASSERT`. Point a
new integration at it before assuming your own linking approach is correct.
