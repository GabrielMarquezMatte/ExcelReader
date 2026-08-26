# ExcelReader.NativeSmoke

A real C consumer of `src/ExcelReader.Native/include/excelreader.h`, run in CI on Windows, Linux and
macOS (`.github/workflows/python.yml`). Supersedes the old untracked root-level `teste.c`.

Two layers, both in `smoke.c`:

1. `_STATIC_ASSERT`s on every ABI struct's `sizeof`/`offsetof`, catching a header layout change on
   whatever compiler builds this file.
2. Runtime assertions driving every export against `RealExcel.xlsb`, catching a mismatch between the
   header and the actual C# implementation (`Exports.cs`/`NativeApi*.cs`) — a layout mismatch there
   produces garbage values, and these assertions fail on the values.

See `smoke.c`'s top comment for why the library is loaded dynamically (`LoadLibrary`/`dlopen`)
instead of linked at build time.

## Build and run

```
python python/scripts/build_native.py    # publishes the library this test loads
cmake -S tests/ExcelReader.NativeSmoke -B build/nativesmoke
cmake --build build/nativesmoke --config Release
ctest --test-dir build/nativesmoke --output-on-failure
```

Pass `-DEXCELREADER_LIB_PATH=<path>` / `-DEXCELREADER_FIXTURE_PATH=<path>` to point at a different
library or fixture; otherwise both default to the standard locations `build_native.py` and the
repository already use.

**Windows without Visual Studio installed:** CMake's default generator needs `nmake`, which only
ships with Visual Studio. Pick MinGW's generator explicitly instead (this is what CI's
`windows-latest` runner does not need, since it has MSVC, but a plain dev machine with only the
.NET SDK and MinGW/MSYS2 installed does):

```
cmake -S tests/ExcelReader.NativeSmoke -B build/nativesmoke -G "MinGW Makefiles"
```
