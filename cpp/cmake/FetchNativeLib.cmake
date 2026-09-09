# Downloads the ExcelReader.Native shared library matching the host platform from a GitHub Release
# asset (published by the publish-native-assets job in .github/workflows/release.yml), and defines
# IMPORTED target xl::native pointing at it.
#
# On Windows, NativeAOT's publish output ships no import library (.lib) - direct linking needs one
# regardless of compiler (MSVC or MinGW), unlike ELF/Mach-O shared libraries. This function
# generates one from src/ExcelReader.Native/include/excelreader.def, the small hand-maintained
# list of symbols this package's phase-1 API actually calls.
#
# Override EXCELREADER_NATIVE_LIB to point at a binary you built locally (e.g. via
# `dotnet publish src/ExcelReader.Native -r win-x64`) instead of downloading a release asset -
# useful for local development and for CI, which builds the native lib fresh per PR (see
# .github/workflows/native-bindings.yml) rather than depending on a tag already being released.
function(excelreader_fetch_native_lib)
    if(WIN32)
        set(_os "win")
        set(_ext "dll")
    elseif(APPLE)
        set(_os "osx")
        set(_ext "dylib")
    else()
        set(_os "linux")
        set(_ext "so")
    endif()

    if(CMAKE_SYSTEM_PROCESSOR MATCHES "^(arm64|aarch64|ARM64)$")
        set(_arch "arm64")
    else()
        set(_arch "x64")
    endif()

    set(_asset_name "excelreader-native-${_os}-${_arch}.${_ext}")
    set(_cache_dir "${CMAKE_BINARY_DIR}/excelreader-native")
    file(MAKE_DIRECTORY "${_cache_dir}")

    if(DEFINED ENV{EXCELREADER_NATIVE_LIB})
        set(_lib_path "$ENV{EXCELREADER_NATIVE_LIB}")
        if(NOT EXISTS "${_lib_path}")
            message(FATAL_ERROR "EXCELREADER_NATIVE_LIB=${_lib_path} does not exist")
        endif()
        # A relative path resolves against the current working directory here (configure time),
        # but is later baked verbatim into IMPORTED_LOCATION/IMPORTED_IMPLIB and re-resolved
        # against the build directory - absolutize it now so it survives that. A no-op if
        # _lib_path is already absolute.
        get_filename_component(_lib_path "${_lib_path}" ABSOLUTE BASE_DIR "${CMAKE_BINARY_DIR}")
    else()
        set(_lib_path "${_cache_dir}/${_asset_name}")
        if(NOT EXISTS "${_lib_path}")
            set(_url "https://github.com/GabrielMarquezMatte/ExcelReader/releases/download/${EXCELREADER_VERSION}/${_asset_name}")
            message(STATUS "Downloading ExcelReader.Native: ${_url}")
            file(DOWNLOAD "${_url}" "${_lib_path}" STATUS _status)
            list(GET _status 0 _code)
            if(NOT _code EQUAL 0)
                file(REMOVE "${_lib_path}")
                list(GET _status 1 _message)
                message(FATAL_ERROR "Failed to download ${_url}: ${_message}")
            endif()
        endif()
    endif()

    add_library(excelreader_native SHARED IMPORTED GLOBAL)
    set_target_properties(excelreader_native PROPERTIES IMPORTED_LOCATION "${_lib_path}")

    if(WIN32)
        set(_src_def_file "${CMAKE_CURRENT_SOURCE_DIR}/include/xl/excelreader.def")
        set(_implib_path "${_cache_dir}/excelreader_native.lib")

        # The checked-in .def has no LIBRARY statement, so neither lib.exe nor dlltool would
        # otherwise know what DLL name to bake into the import descriptors of the generated
        # import lib. That name MUST match the actual basename of the binary at _lib_path
        # (which varies: a release asset like excelreader-native-win-x64.dll, or whatever
        # EXCELREADER_NATIVE_LIB points at locally, e.g. ExcelReader.Native.dll) - not a fixed
        # name - or the loader will look for a DLL that doesn't exist. So generate a temporary
        # .def with an explicit LIBRARY line naming the real file, and feed that to both tools.
        get_filename_component(_dll_basename "${_lib_path}" NAME)
        set(_def_file "${_cache_dir}/excelreader.generated.def")
        file(READ "${_src_def_file}" _def_contents)
        file(WRITE "${_def_file}" "LIBRARY ${_dll_basename}\n${_def_contents}")

        if(NOT EXISTS "${_implib_path}" OR "${_def_file}" IS_NEWER_THAN "${_implib_path}")
            if(MSVC)
                find_program(_lib_exe NAMES lib)
                if(NOT _lib_exe)
                    message(FATAL_ERROR "lib.exe not found - run from a Visual Studio developer prompt")
                endif()
                execute_process(
                    COMMAND "${_lib_exe}" "/def:${_def_file}" "/out:${_implib_path}" "/machine:${_arch}"
                    WORKING_DIRECTORY "${_cache_dir}"
                    RESULT_VARIABLE _lib_result)
                if(NOT _lib_result EQUAL 0)
                    message(FATAL_ERROR "lib.exe /def failed generating the import library")
                endif()
            else()
                find_program(_dlltool NAMES dlltool)
                if(NOT _dlltool)
                    message(FATAL_ERROR "dlltool not found - required to link against ExcelReader.Native.dll under MinGW")
                endif()
                execute_process(
                    COMMAND "${_dlltool}" "-d" "${_def_file}" "-l" "${_implib_path}"
                    RESULT_VARIABLE _dlltool_result)
                if(NOT _dlltool_result EQUAL 0)
                    message(FATAL_ERROR "dlltool failed generating the import library")
                endif()
            endif()
        endif()
        set_target_properties(excelreader_native PROPERTIES IMPORTED_IMPLIB "${_implib_path}")
    endif()

    add_library(xl::native ALIAS excelreader_native)
endfunction()
