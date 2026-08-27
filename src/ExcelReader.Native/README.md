# ExcelReader.Native

NativeAOT shared library exposing ExcelReader's readers over a C ABI, so non-.NET languages
(Python, C, C++, Go, Node) can read XLSX, XLSB, XLS and CSV without a .NET runtime.

- ABI reference and full contract: `include/excelreader.h` (doc comments cover every export)
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

Reading exports: `xl_open_file`, `xl_open_file_ex`, `xl_open_memory`, `xl_open_memory_ex`, `xl_close`,
`xl_sheet_count`, `xl_sheet_name`, `xl_sheet_name_at`, `xl_move_to_sheet`, `xl_is_date1904`,
`xl_next_row`, `xl_read_all_blob`, `xl_read_all_decoded`, `xl_free_rows`, `xl_parse_typed`,
`xl_free_table`, `xl_infer_schema`, `xl_free_schema`, `xl_last_error`, `xl_last_error_ptr`,
`xl_parse_arrow`.

Writing has two layers. The one-shot export, `xl_write_typed` (plus its in-memory twin
`xl_write_typed_to_memory`), takes a whole `xl_table` and writes it in a single call; it takes an
`xl_write_options*` that follows the exact same `struct_size` contract as `xl_open_options`: the
caller sets `options->struct_size = sizeof(xl_write_options)` before the call, and a mismatched value
is rejected with `XL_INVALID_ARGUMENT` before anything else is inspected. It is single-sheet,
whole-table-in-memory, no styling beyond the temporal number formats — see `include/excelreader.h` for
the full contract.

The streaming alternative is `xl_writer_handle`: one sheet and one row open at a time, written
directly as each call arrives instead of building an `xl_table` up front. Call order is
`xl_open_write_handle`/`xl_open_write_handle_to_memory`, then per sheet `xl_start_sheet`..`xl_end_sheet`,
each containing `xl_start_row`..`xl_end_row` with one `xl_write_string`/`xl_write_int64`/
`xl_write_float64`/`xl_write_bool`/`xl_write_date`/`xl_write_time`/`xl_write_timestamp`/`xl_write_null`
call per cell, then `xl_close_write_handle`. `xl_write_handle_bytes` reads back a memory-backed
handle's bytes so far without releasing it. Both write paths release their buffers with
`xl_free_buffer` (`xl_table`/`xl_inferred_schema` keep their own `xl_free_table`/`xl_free_schema`).

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

## Encrypted workbooks

`xl_open_options` grew a `password`/`password_len` pair (UTF-8 bytes, not NUL-terminated) to open a
password-protected OOXML workbook (.xlsx/.xlsb/.xlsm) through `xl_open_file_ex`/`xl_open_memory_ex`.
Omitting it for an encrypted file returns `XL_STATUS_PASSWORD_REQUIRED`; a wrong password returns
`XL_STATUS_PASSWORD_INCORRECT`. Both are new status codes, and `xl_last_error`/`xl_last_error_ptr`
carry the human-readable detail either way:

```c
xl_open_options options = {0};
options.struct_size = sizeof(xl_open_options);
options.password = (const uint8_t*)"hunter2";
options.password_len = 7;

xl_workbook* handle;
int32_t status = xl_open_file_ex(path, path_len, XL_FORMAT_AUTO, &options, &handle);
if (status == XL_STATUS_PASSWORD_REQUIRED || status == XL_STATUS_PASSWORD_INCORRECT) {
    /* prompt again */
}
```

This is why `XL_ABI_VERSION` moved from 3 to 4: `password`/`password_len` are new fields at the end
of `xl_open_options`, so a caller built against the old, smaller struct passes the old, smaller
`sizeof(xl_open_options)` as `struct_size` — the mismatch is rejected outright with
`XL_INVALID_ARGUMENT` instead of the library reading two garbage fields past the end of the caller's
allocation. Check `xl_abi_version()` against `XL_ABI_VERSION` (see above) and rebuild against the
current header rather than only relying on the `struct_size` check to catch it.
