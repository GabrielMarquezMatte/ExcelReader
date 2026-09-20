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

# include/xl is a git symlink to the real header directory. Git on Windows checks a symlink out as a
# text file unless core.symlinks is on, which it is not by default - and FetchContent clones with
# whatever the consumer's git does. Copy the headers into the build tree in that case, so
# `#include <xl/excelreader.hpp>` and the .def below resolve either way.
set(_headers_dir "${CMAKE_CURRENT_SOURCE_DIR}/../src/ExcelReader.Native/include")
if(IS_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}/include/xl")
    set(_include_dir "${CMAKE_CURRENT_SOURCE_DIR}/include")
elseif(IS_DIRECTORY "${_headers_dir}")
    set(_include_dir "${CMAKE_CURRENT_BINARY_DIR}/include")
    file(COPY "${_headers_dir}/" DESTINATION "${_include_dir}/xl")
else()
    message(FATAL_ERROR
        "ExcelReader headers not found: neither ${CMAKE_CURRENT_SOURCE_DIR}/include/xl nor ${_headers_dir} is a directory")
endif()
set(EXCELREADER_DEF_FILE "${_include_dir}/xl/excelreader.def")

add_library(xl_excelreader INTERFACE)
target_include_directories(xl_excelreader INTERFACE "${_include_dir}")
target_compile_features(xl_excelreader INTERFACE cxx_std_23)
add_library(xl::excelreader ALIAS xl_excelreader)

include("${CMAKE_CURRENT_SOURCE_DIR}/cmake/FetchNativeLib.cmake")
include("${CMAKE_CURRENT_SOURCE_DIR}/cmake/CopyNativeLibrary.cmake")
excelreader_fetch_native_lib()
target_link_libraries(xl_excelreader INTERFACE xl::native)

if(CMAKE_SOURCE_DIR STREQUAL PROJECT_SOURCE_DIR)
    set(_install_default ON)
else()
    set(_install_default OFF)
endif()
option(EXCELREADER_INSTALL "Generate install rules and a find_package(excelreader) config package" ${_install_default})

if(EXCELREADER_INSTALL)
    include(GNUInstallDirs)
    include(CMakePackageConfigHelpers)

    get_target_property(_native_lib excelreader_native IMPORTED_LOCATION)
    get_filename_component(EXCELREADER_INSTALLED_NATIVE_NAME "${_native_lib}" NAME)
    if(WIN32)
        set(EXCELREADER_INSTALLED_NATIVE_DIR "${CMAKE_INSTALL_BINDIR}")
    else()
        set(EXCELREADER_INSTALLED_NATIVE_DIR "${CMAKE_INSTALL_LIBDIR}")
    endif()
    set(EXCELREADER_INSTALLED_IMPLIB_DIR "${CMAKE_INSTALL_LIBDIR}")
    set(EXCELREADER_INSTALLED_INCLUDE_DIR "${CMAKE_INSTALL_INCLUDEDIR}")

    install(DIRECTORY "${_headers_dir}/" DESTINATION "${CMAKE_INSTALL_INCLUDEDIR}/xl")
    install(FILES "${_native_lib}" DESTINATION "${EXCELREADER_INSTALLED_NATIVE_DIR}")
    if(WIN32)
        get_target_property(_native_implib excelreader_native IMPORTED_IMPLIB)
        get_filename_component(EXCELREADER_INSTALLED_IMPLIB_NAME "${_native_implib}" NAME)
        install(FILES "${_native_implib}" DESTINATION "${EXCELREADER_INSTALLED_IMPLIB_DIR}")
    endif()

    set(_config_dir "${CMAKE_INSTALL_LIBDIR}/cmake/excelreader")
    configure_package_config_file(
        "${CMAKE_CURRENT_SOURCE_DIR}/cmake/excelreaderConfig.cmake.in"
        "${CMAKE_CURRENT_BINARY_DIR}/excelreaderConfig.cmake"
        INSTALL_DESTINATION "${_config_dir}"
        PATH_VARS
            EXCELREADER_INSTALLED_NATIVE_DIR
            EXCELREADER_INSTALLED_IMPLIB_DIR
            EXCELREADER_INSTALLED_INCLUDE_DIR)

    # EXCELREADER_VERSION carries the release tag ("v3.0.2"); write_basic_package_version_file
    # rejects anything that is not bare digits and dots.
    string(REGEX REPLACE "^v" "" _pkg_version "${EXCELREADER_VERSION}")
    if(NOT _pkg_version MATCHES "^[0-9]+(\\.[0-9]+)*$")
        set(_pkg_version "0.0.0")
    endif()
    write_basic_package_version_file(
        "${CMAKE_CURRENT_BINARY_DIR}/excelreaderConfigVersion.cmake"
        VERSION "${_pkg_version}"
        COMPATIBILITY SameMajorVersion)

    install(FILES
        "${CMAKE_CURRENT_BINARY_DIR}/excelreaderConfig.cmake"
        "${CMAKE_CURRENT_BINARY_DIR}/excelreaderConfigVersion.cmake"
        "${CMAKE_CURRENT_SOURCE_DIR}/cmake/CopyNativeLibrary.cmake"
        DESTINATION "${_config_dir}")
endif()

option(EXCELREADER_BUILD_TESTS "Build the excelreader C++ smoke test" OFF)
if(EXCELREADER_BUILD_TESTS)
    enable_testing()
    add_subdirectory(tests)
endif()

option(EXCELREADER_BUILD_BENCHMARKS "Build the excelreader C++ Google Benchmark suite" OFF)
if(EXCELREADER_BUILD_BENCHMARKS)
    add_subdirectory(benchmarks)
endif()
