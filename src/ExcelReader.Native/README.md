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
