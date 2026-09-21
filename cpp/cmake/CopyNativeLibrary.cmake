# Copies the native shared library next to an executable after it links.
#
# The library is loaded by name at startup, and on Windows a DLL that lives neither beside the
# executable nor on PATH is not found - the process exits with 0xc0000135 (STATUS_DLL_NOT_FOUND)
# before main runs. Everything this package builds calls this, and consumers should too.
#
# A no-op on Linux and macOS, where the rpath baked in at link time already resolves the library.
function(excelreader_copy_native_library target)
    if(NOT WIN32)
        return()
    endif()
    if(NOT TARGET ${target})
        message(FATAL_ERROR "excelreader_copy_native_library: '${target}' is not a target")
    endif()
    add_custom_command(TARGET ${target} POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_FILE:xl::native>" "$<TARGET_FILE_DIR:${target}>"
        COMMENT "Copying ExcelReader.Native next to ${target}")
endfunction()
