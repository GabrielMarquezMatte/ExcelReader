cmake_minimum_required(VERSION 3.16)

# Consumes the installed package the way a downstream project does - find_package() against an
# install prefix, not add_subdirectory() against this repo. Configured as its own project so a
# config file that only works inside the build tree fails here.
project(excelreader_package_consumer LANGUAGES CXX)

find_package(excelreader REQUIRED)

set(EXCELREADER_FIXTURE "" CACHE FILEPATH "Workbook the consumer reads")

add_executable(package_consumer main.cpp)
target_link_libraries(package_consumer PRIVATE xl::excelreader)
target_compile_definitions(package_consumer PRIVATE EXCELREADER_FIXTURE_PATH="${EXCELREADER_FIXTURE}")

if(WIN32)
    add_custom_command(TARGET package_consumer POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_FILE:xl::native>" "$<TARGET_FILE_DIR:package_consumer>")
endif()

enable_testing()
add_test(NAME package_consumer COMMAND package_consumer)
