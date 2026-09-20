add_executable(excelreader_cpp_smoke smoke.cpp)
target_link_libraries(excelreader_cpp_smoke PRIVATE xl::excelreader)

set(EXCELREADER_FIXTURE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/../../RealExcel.xlsb"
    CACHE FILEPATH "Path to the RealExcel.xlsb fixture the C++ smoke test reads")
set(EXCELREADER_ENCRYPTED_FIXTURE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/../../tests/ExcelReader.Tests/data/encrypted"
    CACHE PATH "Directory containing the encrypted-workbook fixtures the C++ smoke test reads")
target_compile_definitions(excelreader_cpp_smoke PRIVATE
    EXCELREADER_FIXTURE_PATH="${EXCELREADER_FIXTURE_PATH}"
    EXCELREADER_ENCRYPTED_FIXTURE_DIR="${EXCELREADER_ENCRYPTED_FIXTURE_DIR}")

add_test(NAME excelreader_cpp_smoke COMMAND excelreader_cpp_smoke)
excelreader_copy_native_library(excelreader_cpp_smoke)

add_executable(excelreader_cpp_write write.cpp)
target_link_libraries(excelreader_cpp_write PRIVATE xl::excelreader)

add_test(NAME excelreader_cpp_write COMMAND excelreader_cpp_write)
excelreader_copy_native_library(excelreader_cpp_write)

add_executable(excelreader_cpp_arrow arrow.cpp)
target_link_libraries(excelreader_cpp_arrow PRIVATE xl::excelreader)
target_compile_definitions(excelreader_cpp_arrow PRIVATE
    EXCELREADER_FIXTURE_PATH="${EXCELREADER_FIXTURE_PATH}")

add_test(NAME excelreader_cpp_arrow COMMAND excelreader_cpp_arrow)
excelreader_copy_native_library(excelreader_cpp_arrow)

add_executable(excelreader_cpp_rows rows.cpp)
target_link_libraries(excelreader_cpp_rows PRIVATE xl::excelreader)

# The existing targets read RealExcel.xlsb; the row tests need the .xlsx alongside it.
set(EXCELREADER_XLSX_FIXTURE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/../../RealExcel.xlsx"
    CACHE FILEPATH "Path to the RealExcel.xlsx fixture the C++ row tests read")
target_compile_definitions(excelreader_cpp_rows PRIVATE
    EXCELREADER_XLSX_FIXTURE_PATH="${EXCELREADER_XLSX_FIXTURE_PATH}")

add_test(NAME excelreader_cpp_rows COMMAND excelreader_cpp_rows)
excelreader_copy_native_library(excelreader_cpp_rows)

add_executable(excelreader_cpp_batch batch.cpp)
target_link_libraries(excelreader_cpp_batch PRIVATE xl::excelreader)

# Reads the same RealExcel.xlsb the smoke test does; that cache variable is set above.
target_compile_definitions(excelreader_cpp_batch PRIVATE
    EXCELREADER_FIXTURE_PATH="${EXCELREADER_FIXTURE_PATH}")

add_test(NAME excelreader_cpp_batch COMMAND excelreader_cpp_batch)
excelreader_copy_native_library(excelreader_cpp_batch)

add_executable(excelreader_cpp_aggregate aggregate.cpp)
target_link_libraries(excelreader_cpp_aggregate PRIVATE xl::excelreader)

add_test(NAME excelreader_cpp_aggregate COMMAND excelreader_cpp_aggregate)
excelreader_copy_native_library(excelreader_cpp_aggregate)