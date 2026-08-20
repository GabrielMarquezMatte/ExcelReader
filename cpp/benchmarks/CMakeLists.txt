include(FetchContent)

# Google Benchmark pulls in its own gtest for BENCHMARK_ENABLE_TESTING; this package only needs
# the benchmark runner itself.
set(BENCHMARK_ENABLE_TESTING OFF CACHE BOOL "" FORCE)
set(BENCHMARK_ENABLE_INSTALL OFF CACHE BOOL "" FORCE)
set(BENCHMARK_ENABLE_WERROR OFF CACHE BOOL "" FORCE)

FetchContent_Declare(
    googlebenchmark
    GIT_REPOSITORY https://github.com/google/benchmark.git
    GIT_TAG v1.9.1
)
FetchContent_MakeAvailable(googlebenchmark)

add_executable(excelreader_cpp_benchmarks benchmark_workbook.cpp)
target_link_libraries(excelreader_cpp_benchmarks PRIVATE xl::excelreader benchmark::benchmark_main)

set(EXCELREADER_FIXTURE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/../../RealExcel.xlsb"
    CACHE FILEPATH "Path to the RealExcel.xlsb fixture the C++ benchmarks read")
set(EXCELREADER_LARGE_FIXTURE_PATH
    "${CMAKE_CURRENT_SOURCE_DIR}/../../tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsb"
    CACHE FILEPATH "Path to the 65K-row fixture the C++ benchmarks read (shared with the .NET benchmark suite)")
target_compile_definitions(excelreader_cpp_benchmarks PRIVATE
    EXCELREADER_FIXTURE_PATH="${EXCELREADER_FIXTURE_PATH}"
    EXCELREADER_LARGE_FIXTURE_PATH="${EXCELREADER_LARGE_FIXTURE_PATH}")

if(WIN32)
    # Same rationale as tests/CMakeLists.txt: the native DLL has no import-time PATH entry.
    add_custom_command(TARGET excelreader_cpp_benchmarks POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_PROPERTY:excelreader_native,IMPORTED_LOCATION>"
            "$<TARGET_FILE_DIR:excelreader_cpp_benchmarks>")
endif()

add_executable(excelreader_cpp_write_benchmarks benchmark_write.cpp)
target_link_libraries(excelreader_cpp_write_benchmarks PRIVATE xl::excelreader benchmark::benchmark_main)
target_compile_definitions(excelreader_cpp_write_benchmarks PRIVATE
    EXCELREADER_LARGE_FIXTURE_PATH="${EXCELREADER_LARGE_FIXTURE_PATH}")

if(WIN32)
    add_custom_command(TARGET excelreader_cpp_write_benchmarks POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_PROPERTY:excelreader_native,IMPORTED_LOCATION>"
            "$<TARGET_FILE_DIR:excelreader_cpp_write_benchmarks>")
endif()

# --- Comparison benchmark against xlnt (https://github.com/tfussell/xlnt) and xlsxio
# (https://github.com/brechtsanders/xlsxio) - separate opt-in flag because both are much heavier
# FetchContent builds (their own zip/xml stacks) than the suite above.
option(EXCELREADER_BUILD_BENCHMARKS_COMPARE "Build the excelreader-vs-xlnt-vs-xlsxio comparison benchmark" OFF)
if(EXCELREADER_BUILD_BENCHMARKS_COMPARE)
    set(STATIC ON CACHE BOOL "" FORCE)
    set(STATIC_CRT OFF CACHE BOOL "" FORCE)
    set(TESTS OFF CACHE BOOL "" FORCE)
    set(SAMPLES OFF CACHE BOOL "" FORCE)
    set(PYTHON OFF CACHE BOOL "" FORCE)
    set(BENCHMARKS OFF CACHE BOOL "" FORCE)

    FetchContent_Declare(
        xlnt
        GIT_REPOSITORY https://github.com/tfussell/xlnt.git
        GIT_TAG v1.5.0
    )
    FetchContent_MakeAvailable(xlnt)

    # xlsxio (https://github.com/brechtsanders/xlsxio) - its own CMakeLists.txt requires an
    # already-installed expat + Minizip (FIND_PACKAGE(... REQUIRED), no FetchContent-friendly
    # fallback), so rather than add_subdirectory-ing it (which would hit those REQUIRED checks),
    # this fetches expat and minizip-ng itself and compiles xlsxio's two read-side .c files
    # directly against them.
    set(EXPAT_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
    set(EXPAT_BUILD_TESTS OFF CACHE BOOL "" FORCE)
    set(EXPAT_BUILD_TOOLS OFF CACHE BOOL "" FORCE)
    set(EXPAT_SHARED_LIBS OFF CACHE BOOL "" FORCE)
    FetchContent_Declare(
        expat
        GIT_REPOSITORY https://github.com/libexpat/libexpat.git
        GIT_TAG R_2_7_5
        SOURCE_SUBDIR expat
    )
    FetchContent_MakeAvailable(expat)

    # minizip-ng's MZ_COMPAT layer provides the classic minizip unzip.h API xlsxio_read.c expects.
    # Only what xlsxio actually needs (zlib decompression, no encryption/other codecs) is enabled.
    set(MZ_COMPAT ON CACHE BOOL "" FORCE)
    set(MZ_ZLIB ON CACHE BOOL "" FORCE)
    set(MZ_ZLIB_FLAVOR "zlib" CACHE STRING "" FORCE)
    # minizip-ng's own MZ_FETCH_LIBS path clones zlib-ng's repo unconditionally (even for the
    # "zlib" flavor, via its stable/compat branch), and that repo's test fixtures have paths too
    # long for Windows' default path limit once fetched into a nested build directory. Its
    # clone_repo() macro honors a pre-set ZLIB_REPOSITORY/ZLIB_TAG though, so point it at the real,
    # much smaller madler/zlib instead - it produces the same zlib/zlibstatic targets clone_repo()
    # expects.
    set(ZLIB_REPOSITORY "https://github.com/madler/zlib.git" CACHE STRING "" FORCE)
    set(ZLIB_TAG "v1.3.1" CACHE STRING "" FORCE)
    # minizip-ng's install(EXPORT ...) step refuses to export the fetched zlibstatic target
    # because its include path lives under this build tree - this is an in-tree add_subdirectory
    # consumed directly via MINIZIP::minizip, so nothing here is ever installed anyway.
    set(SKIP_INSTALL_ALL ON CACHE BOOL "" FORCE)
    set(MZ_BZIP2 OFF CACHE BOOL "" FORCE)
    set(MZ_LZMA OFF CACHE BOOL "" FORCE)
    set(MZ_PPMD OFF CACHE BOOL "" FORCE)
    set(MZ_ZSTD OFF CACHE BOOL "" FORCE)
    set(MZ_PKCRYPT OFF CACHE BOOL "" FORCE)
    set(MZ_WZAES OFF CACHE BOOL "" FORCE)
    set(MZ_ICONV OFF CACHE BOOL "" FORCE)
    set(MZ_BUILD_TESTS OFF CACHE BOOL "" FORCE)
    set(MZ_BUILD_UNIT_TESTS OFF CACHE BOOL "" FORCE)
    set(MZ_FETCH_LIBS ON CACHE BOOL "" FORCE)
    FetchContent_Declare(
        minizipng
        GIT_REPOSITORY https://github.com/zlib-ng/minizip-ng.git
        GIT_TAG 4.2.2
    )
    FetchContent_MakeAvailable(minizipng)

    # xlsxio_read.c does `#include <minizip/unzip.h>`. minizip-ng only installs its compat headers
    # under that exact path name at `cmake --install` time - this is a from-source FetchContent
    # build with no install step, so recreate that one directory locally instead.
    set(_xlsxio_minizip_shim "${CMAKE_CURRENT_BINARY_DIR}/xlsxio-minizip-shim")
    file(MAKE_DIRECTORY "${_xlsxio_minizip_shim}/minizip")
    file(COPY "${minizipng_SOURCE_DIR}/compat/" DESTINATION "${_xlsxio_minizip_shim}/minizip")

    # FetchContent_Populate (not MakeAvailable) deliberately skips add_subdirectory - xlsxio's own
    # CMakeLists.txt is never configured, only its source is downloaded.
    if(POLICY CMP0169)
        cmake_policy(SET CMP0169 OLD)
    endif()
    FetchContent_Declare(
        xlsxio
        GIT_REPOSITORY https://github.com/brechtsanders/xlsxio.git
        GIT_TAG 0.2.36
    )
    FetchContent_Populate(xlsxio)

    add_library(xlsxio_read STATIC
        "${xlsxio_SOURCE_DIR}/lib/xlsxio_read.c"
        "${xlsxio_SOURCE_DIR}/lib/xlsxio_read_sharedstrings.c")
    target_include_directories(xlsxio_read
        PUBLIC "${xlsxio_SOURCE_DIR}/include"
        PRIVATE "${xlsxio_SOURCE_DIR}/lib" "${_xlsxio_minizip_shim}")
    target_compile_definitions(xlsxio_read
        PRIVATE BUILD_XLSXIO USE_MINIZIP
        PUBLIC BUILD_XLSXIO_STATIC)
    target_link_libraries(xlsxio_read PRIVATE MINIZIP::minizip expat::expat)

    # xlsxio's write side is a separate translation unit from its read side, with no shared state -
    # it needs minizip's zip.h (same compat shim as above) but not expat, since writing parses no
    # XML.
    add_library(xlsxio_write STATIC "${xlsxio_SOURCE_DIR}/lib/xlsxio_write.c")
    target_include_directories(xlsxio_write
        PUBLIC "${xlsxio_SOURCE_DIR}/include"
        PRIVATE "${xlsxio_SOURCE_DIR}/lib" "${_xlsxio_minizip_shim}")
    target_compile_definitions(xlsxio_write
        PRIVATE BUILD_XLSXIO USE_MINIZIP
        PUBLIC BUILD_XLSXIO_STATIC)
    target_link_libraries(xlsxio_write PRIVATE MINIZIP::minizip)

    set(EXCELREADER_XLSX_FIXTURE_PATH
        "${CMAKE_CURRENT_SOURCE_DIR}/../../tests/ExcelReader.Benchmarks/Data/65K_Records_Data.xlsx"
        CACHE FILEPATH "Path to the 65K-row xlsx fixture the excelreader-vs-xlnt benchmark reads")

    add_executable(excelreader_cpp_compare_benchmarks benchmark_compare.cpp)
    target_link_libraries(excelreader_cpp_compare_benchmarks PRIVATE
        xl::excelreader benchmark::benchmark_main xlnt xlsxio_read)
    target_compile_definitions(excelreader_cpp_compare_benchmarks PRIVATE
        EXCELREADER_XLSX_FIXTURE_PATH="${EXCELREADER_XLSX_FIXTURE_PATH}")

    # Kept a separate executable from the read comparison rather than a second .cpp in the same
    # one: both files define their own FullRow plus an xl::ExcelMapper specialization for it, and
    # two targets keep that unambiguous.
    add_executable(excelreader_cpp_write_compare_benchmarks benchmark_write_compare.cpp)
    target_link_libraries(excelreader_cpp_write_compare_benchmarks PRIVATE
        xl::excelreader benchmark::benchmark_main xlnt xlsxio_write)
    target_compile_definitions(excelreader_cpp_write_compare_benchmarks PRIVATE
        EXCELREADER_XLSX_FIXTURE_PATH="${EXCELREADER_XLSX_FIXTURE_PATH}")

    if(WIN32)
        foreach(_compare_target excelreader_cpp_compare_benchmarks excelreader_cpp_write_compare_benchmarks)
            add_custom_command(TARGET ${_compare_target} POST_BUILD
                COMMAND ${CMAKE_COMMAND} -E copy_if_different
                    "$<TARGET_PROPERTY:excelreader_native,IMPORTED_LOCATION>"
                    "$<TARGET_FILE_DIR:${_compare_target}>")
        endforeach()
    endif()
endif()
