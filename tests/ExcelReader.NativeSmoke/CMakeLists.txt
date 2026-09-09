# Builds and runs smoke.c against the published ExcelReader.Native shared library. See smoke.c's
# top comment for why the library is loaded dynamically instead of linked at build time.
#
# Usage (from the repository root, after `python python/scripts/build_native.py`):
#     cmake -S tests/ExcelReader.NativeSmoke -B build/nativesmoke
#     cmake --build build/nativesmoke --config Release
#     ctest --test-dir build/nativesmoke --output-on-failure
cmake_minimum_required(VERSION 3.16)
project(ExcelReaderNativeSmoke C)

set(EXCELREADER_ROOT "${CMAKE_CURRENT_SOURCE_DIR}/../..")

# python/scripts/build_native.py already publishes the library and copies it to this fixed,
# OS-normalized path as part of the Python packaging workflow - reusing it here means this project
# never has to guess a runtime identifier (win-x64 vs win-arm64, osx-x64 vs osx-arm64, ...).
if(WIN32)
    set(_lib_name "ExcelReader.Native.dll")
elseif(APPLE)
    set(_lib_name "ExcelReader.Native.dylib")
else()
    set(_lib_name "ExcelReader.Native.so")
endif()

set(EXCELREADER_LIB_PATH "${EXCELREADER_ROOT}/python/src/excelreader/_lib/${_lib_name}"
    CACHE FILEPATH "Path to the published ExcelReader.Native shared library")
set(EXCELREADER_FIXTURE_PATH "${EXCELREADER_ROOT}/RealExcel.xlsb"
    CACHE FILEPATH "Path to the RealExcel.xlsb fixture the smoke test reads")

add_executable(excelreader_smoke smoke.c)
target_include_directories(excelreader_smoke PRIVATE "${EXCELREADER_ROOT}/src/ExcelReader.Native/include")
target_compile_definitions(excelreader_smoke PRIVATE
    EXCELREADER_LIB_PATH_DEFAULT="${EXCELREADER_LIB_PATH}"
    EXCELREADER_FIXTURE_PATH_DEFAULT="${EXCELREADER_FIXTURE_PATH}"
)

if(NOT WIN32)
    # dlopen/dlsym live in libdl on Linux; already part of libc on macOS, where CMAKE_DL_LIBS is empty.
    target_link_libraries(excelreader_smoke PRIVATE ${CMAKE_DL_LIBS})
endif()

enable_testing()
add_test(NAME excelreader_smoke COMMAND excelreader_smoke)
