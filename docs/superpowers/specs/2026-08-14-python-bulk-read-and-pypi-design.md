# Python bulk read, date parsing, and PyPI publishing

Date: 2026-08-14

## Goal

Three independent pieces requested for the Python bindings:

1. Excel serial date parsing, exposed in Python only.
2. A bulk row-materialization API that crosses the native/Python FFI boundary
   once per sheet instead of once per row, so callers don't pay a Python-level
   loop of native calls to get every row.
3. Publish the `excelreader` package to PyPI via CI, in sync with the existing
   NuGet release (`.github/workflows/release.yml`).

## 1. Date parsing (Python only)

No native/C ABI change. `Cell.value` keeps returning the raw serial number as
a string for `CellType.DATE` cells (unchanged contract).

Add to `python/src/excelreader/types.py`:

```python
class Cell(NamedTuple):
    column: int
    type: CellType
    value: str

    def as_date(self, date1904: bool = False) -> datetime.date | None:
        """Converts a DATE cell's serial value. Returns None for any other cell type."""
        if self.type is not CellType.DATE:
            return None
        epoch = datetime.date(1904, 1, 1) if date1904 else datetime.date(1899, 12, 30)
        return epoch + datetime.timedelta(days=int(float(self.value)))
```

This mirrors the exact recipe already asserted in
`test_excel_serial_dates_convert_with_the_documented_recipe`. `date1904` is
supplied by the caller from `Workbook.is_date1904` — `Cell` itself has no
handle to look it up.

Test: move the existing recipe assertion in `test_reader.py` to exercise
`Cell.as_date()` directly, plus a case for a non-DATE cell returning `None`.

## 2. Bulk row materialization

### Native (C ABI)

New export, alongside the existing `xl_next_row_decoded` machinery in
`excelreader.h`:

```c
/* A decoded sheet returned by xl_read_all_decoded. Rows remain valid until xl_free_rows. */
typedef struct xl_rows {
    int32_t row_count;
    xl_row* rows;
} xl_rows;

/* Decodes every remaining row of the current sheet in one call, avoiding one native
 * round-trip per row. The caller owns the returned allocation and must call xl_free_rows. */
int32_t xl_read_all_decoded(void* handle, xl_rows* out_rows);

/* Releases a result returned by xl_read_all_decoded and resets it to zero. Safe on a zeroed value. */
void xl_free_rows(xl_rows* rows);
```

`XL_EOF` is not a failure here — the sheet may legitimately have zero
remaining rows; `xl_read_all_decoded` returns `XL_OK` with `row_count == 0` in
that case instead of `XL_EOF`, since (unlike `xl_next_row`) there's no
per-call "keep going" signal to distinguish from a real error.

Implementation (`NativeApi`, new file `NativeApi.Rows.All.cs` alongside
`NativeApi.Rows.cs`): loop internally calling the same row-enumeration path
`NextRow`/`DecodePendingRow` already use, collecting each decoded `NativeRow`
into one `Marshal.AllocHGlobal` array of `NativeRow` (same struct already
used per-row). `xl_free_rows` calls the existing `FreeRow` on each element,
then frees the array.

`Exports.cs` gets two more `[UnmanagedCallersOnly]` entries following the
exact shape of `NextRowDecoded`/`FreeRow`.

### Python

`_native.py`: bind `xl_read_all_decoded` / `xl_free_rows` (new `NativeRows`
ctypes Structure mirroring `xl_rows`, reusing whatever ctypes Structure
already backs `xl_row`/`xl_row_cell` for `xl_next_row_decoded` — check
whether that binding already exists from the prior `xl_next_row_decoded`
native addition; add it if it doesn't).

`reader.py`: `Workbook.read_all(self) -> list[list[Cell]]`. One native call,
one Python-side pass decoding the returned `xl_rows` into
`list[list[Cell]]`, then `xl_free_rows` in a `finally`. No Python loop of
native calls — the sheet-wide loop lives in `NativeApi`, not in
`reader.py`.

Out of scope (explicitly deferred, noted so it isn't lost): pandas/polars
DataFrame output. Today's return shape is `list[list[Cell]]`, matching what
was requested; a columnar/typed extension is future work once this shape is
in place.

Test: a `test_read_all_matches_row_by_row_iteration` asserting
`workbook.read_all()` equals `list(workbook.rows())` for the xlsx fixture,
plus an empty-sheet-at-EOF case (`move_to_sheet` to the end, or call
`read_all` twice) returning `[]` rather than raising.

## 3. Publish to PyPI via CI

### Wheel tagging problem

`_lib/` holds exactly one platform's native binary per build
(`build_native.py` copies only the current OS's library). A wheel built this
way is not pure Python, but hatchling's default wheel tag is `py3-none-any`
— three OS-specific wheels would all get that same filename+tag and collide
on upload (PyPI rejects a re-uploaded filename).

Fix: a hatchling build hook forcing hatchling to infer the real platform tag
from the machine building it. New file `python/hatch_build.py`:

```python
from hatchling.builders.hooks.plugin.interface import BuildHookInterface


class NativeLibraryBuildHook(BuildHookInterface):
    def initialize(self, version, build_data):
        build_data["pure_python"] = False
        build_data["infer_tag"] = True
```

Registered in `pyproject.toml`:

```toml
[tool.hatch.build.targets.wheel.hooks.custom]
path = "hatch_build.py"
```

This is a documented hatchling mechanism (no new dependency) — the wheel
filename becomes e.g. `excelreader-1.2.3-cp312-cp312-win_amd64.whl` /
`...-manylinux...` / `...-macosx...`, one distinct filename per OS.

### Versioning

`pyproject.toml` keeps a static placeholder version. The release job derives
the real version the same way `release.yml` already does for NuGet
(`${GITHUB_REF_NAME#v}`) and writes it into `pyproject.toml` before building
— a one-line `sed`/Python replace, no dynamic-versioning plugin.

### Workflow

Add a `publish-python` job to the existing `.github/workflows/release.yml`
(same `v*` tag, so a single tag publishes both NuGet and PyPI):

```yaml
  publish-python:
    name: Build & Publish to PyPI
    runs-on: ${{ matrix.os }}
    permissions:
      id-token: write
    strategy:
      matrix:
        os: [ubuntu-latest, windows-latest, macos-latest]
    steps:
      - checkout (fetch-depth: 0)
      - setup .NET, setup Python
      - derive version from tag (reuse steps.version pattern)
      - write version into python/pyproject.toml
      - python python/scripts/build_native.py
      - pip install build
      - python -m build --wheel python/ --outdir dist/
      - upload-artifact (per-OS wheel)

  publish-pypi:
    needs: publish-python
    runs-on: ubuntu-latest
    permissions:
      id-token: write
    steps:
      - download all wheel artifacts into dist/
      - build the sdist once (any OS, pure Python + metadata)
      - pypa/gh-action-pypi-publish@release/v1   # Trusted Publishing, no API token
```

Trusted Publishing (OIDC) mirrors the `NuGet/login@v1` pattern already used
for NuGet — configured once on the PyPI project side (pending user setup on
pypi.org), no secret stored in the repo.

### Out of scope

- TestPyPI dry-runs, changelogs, and the `promote-public-api` step (that one
  is C#-specific `PublicAPI.Unshipped`/`Shipped` tracking, not applicable to
  the Python package).

## Testing summary

- `Cell.as_date()`: unit tests in `python/tests/test_reader.py` (or a new
  `test_types.py`).
- `xl_read_all_decoded`/`xl_free_rows`: a native-side test alongside the
  existing `xl_next_row_decoded` tests (mirroring their EOF/error-path
  coverage), plus the Python `test_read_all_matches_row_by_row_iteration`.
- CI: verified by the workflow actually running on the next tag push; no way
  to unit-test a GitHub Actions workflow file itself beyond `actionlint`/a
  dry run.
