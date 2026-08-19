# Multi-name column binding (native bindings) — design

## Context

`ExcelColumnAttribute` in the C# core (`src/ExcelReader.Core/Parser/ExcelColumnAttribute.cs`)
is `[AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]`:
a property can carry it more than once, each occurrence naming an
alternate header spelling. `TypeMapper<T>.Build`
(`src/ExcelReader.Core/Parser/Internal/TypeMapper.cs:82-86`) collects
every name into a `string[]` on `PropertyMap<T>.Names`, and
`TypeMapInfo<T>.BuildLookup`
(`src/ExcelReader.Core/Parser/Internal/TypeMapInfo.cs:143-157`) inserts
them into a `Dictionary<string, HeaderMatch<T>>` via `TryAdd` in
declared order — the first alias that reaches a given header name wins,
giving "first match" semantics for a property with several candidate
column names.

None of the native bindings (Rust, C++, Python) have an equivalent: each
maps exactly one field/column to exactly one header name.

- **Rust** (`rust/excelreader-derive/src/lib.rs`): `#[excel(name = "...")]`
  per field; `excel_name` (lines 80-103) returns the first (only)
  `name = ...` value and the macro emits one `ColumnBinding { name, .. }`
  (`ColumnBinding.name: &'static str`, `rust/excelreader/src/workbook.rs:295-299`).
- **C++** (`src/ExcelReader.Native/include/excelreader.hpp`):
  `make_field(const char *name, T Class::*member)` (lines 476-480) builds
  one `FieldBinding<Class, T>` (lines 468-474) per field, each holding a
  single `column_name`.
- **Python** (`python/src/excelreader/`): no per-field mapper at all —
  callers build a `Sequence[ColumnSpec]` by hand
  (`ColumnSpec.name: str | None`, `types.py:79-90`) and pass it to
  `parse_typed`/`write_typed`. `_native.column_spec_by_name` /
  `_build_specs` (`reader.py:400-409`) turn each `ColumnSpec` into one
  `_native.NativeColumnSpec` with a single `name`.

All three ultimately cross the same C ABI struct, resolved by the same
native core:

```c
/* include/excelreader.h */
typedef struct xl_column_spec {
    const uint8_t* name;   /* header text to match, UTF-8; NULL to match by index instead */
    int32_t name_len;
    int32_t index;
    int32_t type;
    int32_t nullable;
} xl_column_spec;
```

`xl_parse_typed` is *implemented in this repo*, in C# NativeAOT code
(`src/ExcelReader.Native`) — not an external prebuilt library (confirmed
via `cpp/cmake/FetchNativeLib.cmake`: the shared library is a GitHub
Release asset built by this repo's own `dotnet publish`, or a local
build). `NativeApi.Typed.TryResolveColumns` /
`FindHeaderColumn` (`NativeApi.Typed.cs:189-245`) is the actual
name-to-column-index resolution: it scans the header row once and, for
each spec, calls `FindHeaderColumn(header, spec.Name)` — one name in,
one match out. This is the one place "first match among several
candidates" needs to be taught, and every binding (Rust, C++, Python)
inherits the fix for free once it's expressed as an array in the ABI
struct.

This spec adds that: a property/field/column spec may carry an ordered
list of candidate header names, and the first one present in the header
row wins — mirroring the C# semantics, implemented once in the shared
native core.

## Goal

```rust
// Rust
#[excel(name = "Nome", alias = "Nom", alias = "Name")]
nome: String,
```

```cpp
// C++
make_field({"Nome", "Nom", "Name"}, &Row::nome)
```

```python
# Python
ColumnSpec(ColumnType.STRING, name=["Nome", "Nom", "Name"])
```

In every case: at parse time, the native core tries `"Nome"` against the
header row, then `"Nom"`, then `"Name"`, and binds to the first one that
matches (case-insensitive, trimmed — unchanged matching rule). If none
match, the call fails the same way an unmatched single name does today
(`XL_INVALID_ARGUMENT`, `xl_last_error` naming the field).

## ABI change (breaking, ABI version bump)

`xl_column_spec` moves from a single name to an ordered array of
candidate names:

```c
typedef struct xl_column_spec {
    const uint8_t* const* names; /* candidate header texts, UTF-8, in priority order;
                                     NULL (name_count == 0) to match by index instead */
    const int32_t* name_lens;    /* one length per entry in `names` */
    int32_t name_count;
    int32_t index;                /* zero-based column index, used when name_count == 0 */
    int32_t type;                 /* XL_T_* */
    int32_t nullable;
} xl_column_spec;
```

- `XL_ABI_VERSION` (`include/excelreader.h`, mirrored in
  `python/src/excelreader/_native.py`) is bumped by 1 — this is exactly
  the "struct layout changed" case the existing version-check comment
  describes.
- `xl_parse_typed`/`xl_parse_arrow` (read direction): resolution tries
  `names[0]`, then `names[1]`, ... against the header row in order; the
  first match wins. Unmatched → `XL_INVALID_ARGUMENT`, same as today but
  the error message lists every candidate that was tried.
- `xl_write_typed` (write direction) and `xl_infer_schema` (output
  direction) are unaffected in behavior: a write spec is still exactly
  one name (`name_count` must be `1` for a name-based write spec — `0`
  or `>1` is `XL_INVALID_ARGUMENT`), and an inferred column always comes
  back with `name_count` either `0` (index-based) or `1` (a single
  guessed header). Neither gets new capability here; they keep working
  because the struct shape still describes "zero, one, or many names."
- Validation (`NativeApi.Typed.TryValidateArguments`) walks the whole
  `Names` array per spec: any blank entry, or a name-based spec with
  `header_row == 0`, fails the call exactly as a single blank/misplaced
  name does today.

## Native core (`src/ExcelReader.Native`) — source of truth

- `NativeColumnSpecRaw` (`NativeTypedTable.cs:27-34`): replace
  `byte* Name; int NameLen;` with `byte** Names; int* NameLens; int NameCount;`.
- `NativeColumnSpec` (`NativeTypedTable.cs:38-46`): replace
  `string? Name` with `string[] Names` (empty array = index-based,
  mirroring today's `null`).
- `Exports.TryDecodeColumnSpecs` (`Exports.cs:401-421`): decode the
  `NameCount`-length pointer/length arrays into `string[]`, same
  `IsValidNameLength` bound applied per entry.
- `NativeApi.Typed.cs`:
  - `TryValidateArguments` (145-184): "name-based" becomes
    `spec.Names.Length > 0`; the blank-name check and the
    `header_row == 0` check iterate `Names`.
  - `TryResolveColumns`/`FindHeaderColumn` (189-245): for each spec, try
    each `Names[i]` against the header row in order; first match sets
    `columnIndices[i]` and stops. No match after exhausting the array →
    error listing every candidate tried (`no column header matches any of "Nome", "Nom", "Name".`).
  - The error path in `ParseTyped`'s row loop (line 59) that names the
    failing spec (`spec.Name ?? spec.Index...`) becomes
    `spec.Names.Length > 0 ? string.Join(" / ", spec.Names) : spec.Index...`.
- `NativeApi.Write.cs` / write-path validation: a name-based write spec
  must have exactly one entry in `Names` — this is new validation (today
  a write spec structurally can't have more than one name), rejected as
  `XL_INVALID_ARGUMENT` with a message explaining a write column takes
  exactly one name.
- `NativeApi.Schema.cs` (`xl_infer_schema`): unaffected in logic — it
  already produces one name (or none) per guessed column; it now writes
  that into the new `names`/`name_lens`/`name_count` output shape
  (`name_count` 0 or 1).

## C++ (`excelreader.hpp`)

- `xl_column_spec`'s mirrored C++ struct picks up the same
  `names`/`name_lens`/`name_count` fields.
- `FieldBinding<Class, T>` (468-474) stores `std::vector<const char*> column_names`
  instead of a single `column_name`.
- `make_field` gains an overload:
  ```cpp
  template <typename Class, typename T>
  FieldBinding<Class, T> make_field(std::initializer_list<const char*> names, T Class::*member);
  ```
  The existing single-name `make_field(const char *name, T Class::*member)`
  becomes a thin forward to the list overload with one element — existing
  call sites keep compiling unchanged (source-compatible; the ABI change
  is still a breaking rebuild, since the struct layout moved).
- `build_specs` (489-499) marshals each binding's `column_names` vector
  into the new `names`/`name_lens`/`name_count` layout instead of a
  single pointer/length pair.

## Rust (`excelreader` + `excelreader-derive`)

- `ColumnBinding<T>` (`workbook.rs:295-299`): `name: &'static str` becomes
  `names: &'static [&'static str]`.
- Macro syntax: `#[excel(name = "Nome", alias = "Nom", alias = "Name")]` —
  `name` stays required and first in priority order; `alias` is optional
  and repeatable within the same `#[excel(...)]` attribute.
  `excel_name` (80-103) becomes `excel_names`, collecting `name` followed
  by every `alias` into a `Vec<LitStr>` (compile error, as today, if
  `name` is missing).
- `field_binding` (58-78) emits `names: &[#(#names),*]` instead of a
  single `name: #name`.
- `parse_sheet`/`parse_arrow` (`workbook.rs:499+`): building
  `XlColumnSpec` from `ColumnBinding` changes from one
  pointer/length pair to marshaling `binding.names` into a
  `Vec<*const u8>` / `Vec<i32>` pair kept alive for the duration of the
  native call (same lifetime pattern the current single-name version
  already has to respect, just one level deeper).

## Python (`python/src/excelreader/`)

- `_native.py`:
  - `NativeColumnSpec` (70-79): replace `name`/`name_len` fields with
    `names: POINTER(POINTER(c_uint8))`, `name_lens: POINTER(c_int32)`,
    `name_count: c_int32`. `ctypes.c_char_p` doesn't compose into an
    array-of-pointers cleanly with embedded lengths (names aren't
    NUL-terminated on the wire), so `names` moves to explicit
    `uint8_t**` + `name_lens`, matching how `NativeInferredColumnSpec`
    already treats a single name today.
  - `column_spec_by_name` (82-84) becomes `column_spec_by_names(names: Sequence[str], ...)`:
    encodes each candidate to UTF-8, builds a `(POINTER(c_uint8) * n)()`
    array of owned buffers (kept alive via the returned struct's
    `_objects`, same pattern `to_native_write_options` already relies on
    for `sheet_name`) and a matching `(c_int32 * n)()` length array.
    `column_spec_by_name(name: str, ...)` stays as a one-element
    convenience wrapper over it so existing call sites keep working.
  - `NativeInferredColumnSpec` (91-106): same field rename for
    consistency, but the native side only ever fills `name_count` 0 or 1
    for this output-only struct — `_decode_inferred_schema` reads
    `raw.names[0]`/`raw.name_lens[0]` when `name_count == 1`, `None`
    otherwise.
  - `XL_ABI_VERSION` bumped to match the C header.
- `types.py`: `ColumnSpec.name` widens from `str | None` to
  `str | Sequence[str] | None`. A plain `str` keeps meaning exactly what
  it means today (one candidate); a `Sequence[str]` is the new ordered
  multi-candidate form. This is source-compatible for every existing
  caller.
- `reader.py` / `writer.py`:
  - `_build_specs` (400-409): a `str` name goes through
    `column_spec_by_names([spec.name], ...)`; a `Sequence[str]` goes
    through directly. `spec.name is None` (index-based) is unchanged.
  - `writer.py:137-140` (`xl_write_typed` specs): still always exactly
    one name per column (write direction never takes a candidate list);
    continues to call `column_spec_by_name` (the single-name wrapper).
  - `_decode_inferred_schema` (391-397): reads through the renamed
    `NativeInferredColumnSpec` fields, keeps returning `ColumnSpec.name`
    as a plain `str | None` (inference never guesses more than one name
    per column).

## Testing

- **C#**: `tests/ExcelReader.Tests/NativeApiTests.cs` — new cases: a spec
  with 2+ names where only the second/third matches the header (first
  match wins when several would match), a spec where none match
  (message lists every candidate), a write spec with `name_count != 1`
  rejected. Existing single-name cases become 1-element-array cases.
- **C**: `tests/ExcelReader.NativeSmoke/smoke.c` — update spec
  construction to the array layout; add one multi-name smoke case.
- **C++**: `cpp/tests` — a `make_field({...}, member)` case exercising
  second-candidate match; existing single-name `make_field` calls stay
  as-is (still route through the same code path).
- **Rust**: `excelreader-derive`'s `trybuild`/unit tests — a struct field
  with `name` + multiple `alias`es, asserting the emitted `names` slice
  and priority order; `excelreader/tests/parse_typed.rs` — an end-to-end
  case where the sheet's header uses an alias, not the primary `name`.
- **Python**: `python/tests` — `_build_specs` with a `Sequence[str]`
  name resolving to the second candidate; `ColumnSpec(name="X")` (plain
  str) still round-trips; ABI version mismatch still raises with the new
  version bumped.

## Out of scope

- Any change to the C# core (`ExcelColumnAttribute`/`TypeMapper`) —
  already has this feature; nothing to do there.
- A "warn on duplicate header claimed by two properties" analyzer for
  the native bindings, equivalent to the C# source generator's
  `DuplicateHeaderDescriptor`. Not implemented here; collisions resolve
  however they happen to today (first spec in the caller-supplied list
  wins, unchanged), just not diagnosed at compile/macro time.
- Changing the header-matching rule itself (case-insensitive, trimmed) —
  unchanged, just applied per candidate instead of once.
