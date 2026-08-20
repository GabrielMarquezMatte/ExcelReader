add_executable(excelreader_cpp_smoke smoke.cpp)
target_link_libraries(excelreader_cpp_smoke PRIVATE xl::excelreader)

set(EXCELREADER_FIXTURE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/../../RealExcel.xlsb"
    CACHE FILEPATH "Path to the RealExcel.xlsb fixture the C++ smoke test reads")
target_compile_definitions(excelreader_cpp_smoke PRIVATE
    EXCELREADER_FIXTURE_PATH="${EXCELREADER_FIXTURE_PATH}")

add_test(NAME excelreader_cpp_smoke COMMAND excelreader_cpp_smoke)
if(WIN32)
    # The native DLL has no import-time PATH entry; copy it next to the test executable so the
    # loader finds it without requiring the caller to fiddle with PATH.
    add_custom_command(TARGET excelreader_cpp_smoke POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_PROPERTY:excelreader_native,IMPORTED_LOCATION>"
            "$<TARGET_FILE_DIR:excelreader_cpp_smoke>")
endif()
