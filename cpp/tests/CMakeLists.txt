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
if(WIN32)
    # The native DLL has no import-time PATH entry; copy it next to the test executable so the
    # loader finds it without requiring the caller to fiddle with PATH.
    add_custom_command(TARGET excelreader_cpp_smoke POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_PROPERTY:excelreader_native,IMPORTED_LOCATION>"
            "$<TARGET_FILE_DIR:excelreader_cpp_smoke>")
endif()

add_executable(excelreader_cpp_write write.cpp)
target_link_libraries(excelreader_cpp_write PRIVATE xl::excelreader)

add_test(NAME excelreader_cpp_write COMMAND excelreader_cpp_write)
if(WIN32)
    # Same reason as the smoke test: the native DLL has no import-time PATH entry, so copy it next
    # to the test executable for the loader to find.
    add_custom_command(TARGET excelreader_cpp_write POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_PROPERTY:excelreader_native,IMPORTED_LOCATION>"
            "$<TARGET_FILE_DIR:excelreader_cpp_write>")
endif()

add_executable(excelreader_cpp_arrow arrow.cpp)
target_link_libraries(excelreader_cpp_arrow PRIVATE xl::excelreader)
target_compile_definitions(excelreader_cpp_arrow PRIVATE
    EXCELREADER_FIXTURE_PATH="${EXCELREADER_FIXTURE_PATH}")

add_test(NAME excelreader_cpp_arrow COMMAND excelreader_cpp_arrow)
if(WIN32)
    # Same reason as the smoke test: the native DLL has no import-time PATH entry, so copy it next
    # to the test executable for the loader to find.
    add_custom_command(TARGET excelreader_cpp_arrow POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_PROPERTY:excelreader_native,IMPORTED_LOCATION>"
            "$<TARGET_FILE_DIR:excelreader_cpp_arrow>")
endif()

add_executable(excelreader_cpp_rows rows.cpp)
target_link_libraries(excelreader_cpp_rows PRIVATE xl::excelreader)

# The existing targets read RealExcel.xlsb; the row tests need the .xlsx alongside it.
set(EXCELREADER_XLSX_FIXTURE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/../../RealExcel.xlsx"
    CACHE FILEPATH "Path to the RealExcel.xlsx fixture the C++ row tests read")
target_compile_definitions(excelreader_cpp_rows PRIVATE
    EXCELREADER_XLSX_FIXTURE_PATH="${EXCELREADER_XLSX_FIXTURE_PATH}")

add_test(NAME excelreader_cpp_rows COMMAND excelreader_cpp_rows)
if(WIN32)
    # Same reason as the smoke test: the native DLL has no import-time PATH entry, so copy it next
    # to the test executable for the loader to find.
    add_custom_command(TARGET excelreader_cpp_rows POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_PROPERTY:excelreader_native,IMPORTED_LOCATION>"
            "$<TARGET_FILE_DIR:excelreader_cpp_rows>")
endif()