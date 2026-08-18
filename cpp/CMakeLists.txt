cmake_minimum_required(VERSION 3.16)
project(ExcelReader VERSION 0.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 23)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(EXCELREADER_VERSION "" CACHE STRING
    "Release tag whose assets FetchNativeLib.cmake downloads (auto-detected from `git describe` if empty)")
if(NOT EXCELREADER_VERSION)
    find_package(Git QUIET)
    set(_git_tag "")
    if(Git_FOUND)
        execute_process(
            COMMAND ${GIT_EXECUTABLE} describe --tags --exact-match
            WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
            OUTPUT_VARIABLE _git_tag
            OUTPUT_STRIP_TRAILING_WHITESPACE
            ERROR_QUIET
            RESULT_VARIABLE _git_tag_result)
    endif()
    if(_git_tag_result EQUAL 0 AND _git_tag)
        set(EXCELREADER_VERSION "${_git_tag}")
    else()
        set(EXCELREADER_VERSION "v0.0.0")
    endif()
endif()

add_library(xl_excelreader INTERFACE)
target_include_directories(xl_excelreader INTERFACE "${CMAKE_CURRENT_SOURCE_DIR}/include")
target_compile_features(xl_excelreader INTERFACE cxx_std_23)
add_library(xl::excelreader ALIAS xl_excelreader)

include("${CMAKE_CURRENT_SOURCE_DIR}/cmake/FetchNativeLib.cmake")
excelreader_fetch_native_lib()
target_link_libraries(xl_excelreader INTERFACE xl::native)

option(EXCELREADER_BUILD_TESTS "Build the excelreader C++ smoke test" OFF)
if(EXCELREADER_BUILD_TESTS)
    enable_testing()
    add_subdirectory(tests)
endif()
