//! Downloads the ExcelReader.Native shared library matching the build target, and (on
//! windows-msvc/windows-gnu) generates the import library NativeAOT's publish output doesn't ship,
//! from `excelreader.def` (a checked-in copy of the phase-1 .def next to excelreader.h,
//! kept inside this crate directory - see the comment on `PHASE1_DEF_EXPORTS` below - so it's
//! embedded via `include_str!` and survives `cargo package`).
//!
//! Override EXCELREADER_NATIVE_LIB_DIR to point at a directory containing a locally-built
//! ExcelReader.Native.{dll,so,dylib} instead of downloading a release asset - used by CI (see
//! .github/workflows/rust.yml), which builds the native lib fresh per PR rather than depending on a
//! tag already being released, and by local development.

use std::env;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

const REPO: &str = "GabrielMarquezMatte/ExcelReader";

fn main() {
    println!("cargo:rerun-if-env-changed=EXCELREADER_NATIVE_LIB_DIR");
    println!("cargo:rerun-if-changed=excelreader.def");

    let out_dir = PathBuf::from(env::var("OUT_DIR").unwrap());
    let target_os = env::var("CARGO_CFG_TARGET_OS").unwrap();
    let target_env = env::var("CARGO_CFG_TARGET_ENV").unwrap_or_default();
    let target_arch = env::var("CARGO_CFG_TARGET_ARCH").unwrap();

    let (os, ext) = match target_os.as_str() {
        "windows" => ("win", "dll"),
        "macos" => ("osx", "dylib"),
        "linux" => ("linux", "so"),
        other => panic!("unsupported target_os for excelreader: {other}"),
    };
    let arch = match target_arch.as_str() {
        "aarch64" => "arm64",
        _ => "x64",
    };
    let asset_name = format!("excelreader-native-{os}-{arch}.{ext}");

    let (lib_dir, dll_basename) = if let Ok(dir) = env::var("EXCELREADER_NATIVE_LIB_DIR") {
        let dir = PathBuf::from(dir);
        let dll_basename = find_native_lib(&dir, ext);
        (dir, dll_basename)
    } else {
        let dest = out_dir.join(&asset_name);
        if !dest.exists() {
            let version = env::var("CARGO_PKG_VERSION").unwrap();
            let url =
                format!("https://github.com/{REPO}/releases/download/v{version}/{asset_name}");
            download(&url, &dest);
        }
        (out_dir.clone(), asset_name)
    };

    println!("cargo:rustc-link-search=native={}", lib_dir.display());

    if target_os == "windows" {
        let implib = out_dir.join("excelreader_native.lib");
        let generated_def = out_dir.join("excelreader.generated.def");
        let new_def_content = format!("LIBRARY {dll_basename}\n{PHASE1_DEF_EXPORTS}");
        let up_to_date = implib.exists()
            && fs::read_to_string(&generated_def).is_ok_and(|old| old == new_def_content);
        if !up_to_date {
            generate_windows_implib(
                &implib, &generated_def, &new_def_content, &target_env, &target_arch, &out_dir,
                &dll_basename,
            );
        }
        println!("cargo:rustc-link-search=native={}", out_dir.display());
        println!("cargo:rustc-link-lib=dylib=excelreader_native");
        if let Some(profile_dir) = out_dir.ancestors().nth(3) {
            let deps_dir = profile_dir.join("deps");
            fs::create_dir_all(&deps_dir)
                .unwrap_or_else(|e| panic!("failed to create {}: {e}", deps_dir.display()));
            let dest = deps_dir.join(&dll_basename);
            fs::copy(lib_dir.join(&dll_basename), &dest).unwrap_or_else(|e| {
                panic!("failed to copy {dll_basename} to {}: {e}", dest.display())
            });
        }
    } else {
        println!(
            "cargo:rustc-link-arg={}",
            lib_dir.join(&dll_basename).display()
        );
        println!("cargo:rustc-link-arg=-Wl,-rpath,{}", lib_dir.display());
    }
}

/// Scans `dir` for a single file with extension `ext` and returns its basename. Used for the
/// `EXCELREADER_NATIVE_LIB_DIR` override, where the exact filename of the native binary isn't
/// known ahead of time (unlike the download path, where it's always `asset_name`).
fn find_native_lib(dir: &Path, ext: &str) -> String {
    let mut candidates: Vec<String> = fs::read_dir(dir)
        .unwrap_or_else(|e| {
            panic!(
                "failed to read EXCELREADER_NATIVE_LIB_DIR {}: {e}",
                dir.display()
            )
        })
        .filter_map(|entry| entry.ok())
        .filter_map(|entry| {
            let path = entry.path();
            if path.extension().and_then(|e| e.to_str()) == Some(ext) {
                path.file_name()?.to_str().map(String::from)
            } else {
                None
            }
        })
        .collect();
    candidates.sort();
    match candidates.len() {
        0 => panic!(
            "no .{ext} file found in EXCELREADER_NATIVE_LIB_DIR {}",
            dir.display()
        ),
        1 => candidates.remove(0),
        _ => panic!(
            "multiple .{ext} files found in EXCELREADER_NATIVE_LIB_DIR {}: {candidates:?} - expected exactly one",
            dir.display()
        ),
    }
}

fn download(url: &str, dest: &Path) {
    let status = if cfg!(windows) {
        Command::new("powershell")
            .args([
                "-NoProfile",
                "-Command",
                &format!(
                    "Invoke-WebRequest -Uri '{url}' -OutFile '{}'",
                    dest.display()
                ),
            ])
            .status()
    } else {
        Command::new("curl")
            .args(["-fSL", "-o"])
            .arg(dest)
            .arg(url)
            .status()
    };
    match status {
        Ok(s) if s.success() => {}
        Ok(s) => panic!("downloading {url} failed with exit status {s}"),
        Err(e) => panic!("failed to invoke downloader for {url}: {e}"),
    }
}

const PHASE1_DEF_EXPORTS: &str = include_str!("excelreader.def");

fn generate_windows_implib(
    implib: &Path,
    generated_def: &Path,
    def_content: &str,
    target_env: &str,
    arch: &str,
    out_dir: &Path,
    dll_basename: &str,
) {
    fs::write(generated_def, def_content)
        .unwrap_or_else(|e| panic!("failed to write {}: {e}", generated_def.display()));

    if target_env == "msvc" {
        let lib_exe = "lib.exe";
        let status = Command::new(lib_exe)
            .arg(format!("/def:{}", generated_def.display()))
            .arg(format!("/out:{}", implib.display()))
            .arg(format!("/machine:{arch}"))
            .status()
            .unwrap_or_else(|e| {
                panic!("failed to invoke {lib_exe} (run from a VS developer prompt): {e}")
            });
        assert!(status.success(), "{lib_exe} /def failed");
    } else {
        let implib_name = implib
            .file_name()
            .expect("implib path has no file name")
            .to_str()
            .unwrap();
        let status = Command::new("dlltool")
            .current_dir(out_dir)
            .args(["-d", generated_def.to_str().unwrap()])
            .args(["-l", implib_name])
            .args(["-D", dll_basename])
            .status()
            .unwrap_or_else(|e| panic!("failed to invoke dlltool: {e}"));
        assert!(status.success(), "dlltool failed");
    }
    let _ = fs::metadata(implib).expect("import library was not created");
}
