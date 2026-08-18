//! Downloads the ExcelReader.Native shared library matching the build target, and (on
//! windows-msvc/windows-gnu) generates the import library NativeAOT's publish output doesn't ship,
//! from `excelreader-phase1.def` (a checked-in copy of the phase-1 .def next to excelreader.h,
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

    // The actual filename of the native binary that ends up in `lib_dir`. For a downloaded
    // release asset this is always `asset_name`. For the `EXCELREADER_NATIVE_LIB_DIR` override
    // it can be anything (e.g. `ExcelReader.Native.dll` from a local `dotnet publish`), so it
    // must be discovered by scanning the directory rather than assumed.
    let (lib_dir, dll_basename) = if let Ok(dir) = env::var("EXCELREADER_NATIVE_LIB_DIR") {
        let dir = PathBuf::from(dir);
        let dll_basename = find_native_lib(&dir, ext);
        (dir, dll_basename)
    } else {
        let dest = out_dir.join(&asset_name);
        if !dest.exists() {
            let version = env::var("CARGO_PKG_VERSION").unwrap();
            let url = format!(
                "https://github.com/{REPO}/releases/download/v{version}/{asset_name}"
            );
            download(&url, &dest);
        }
        (out_dir.clone(), asset_name)
    };

    println!("cargo:rustc-link-search=native={}", lib_dir.display());

    if target_os == "windows" {
        let implib = out_dir.join("excelreader_native.lib");
        if !implib.exists() {
            generate_windows_implib(
                &implib,
                &target_env,
                &target_arch,
                &out_dir,
                &dll_basename,
            );
        }
        // The import library lives in OUT_DIR, not lib_dir (the DLL's own directory) - without
        // this, the linker can locate the DLL for cargo:rustc-link-search purposes but never finds
        // excelreader_native.lib/.a itself, so any real link step (e.g. `cargo test`) fails.
        println!("cargo:rustc-link-search=native={}", out_dir.display());
        // The import library above is always generated with this fixed name, regardless of what
        // the underlying DLL is actually called (its LIBRARY line handles that indirection), so a
        // plain name-based link works here.
        println!("cargo:rustc-link-lib=dylib=excelreader_native");
    } else {
        // On macOS/Linux there's no import library indirection - rustc/the linker must find the
        // shared object under its own real name. That name is NOT a fixed `libexcelreader_native.*`
        // - it's `asset_name` on the download path (e.g. `excelreader-native-linux-x64.so`) or
        // whatever `find_native_lib` discovered on the `EXCELREADER_NATIVE_LIB_DIR` override path
        // (e.g. `ExcelReader.Native.dylib`), neither of which fits the conventional `lib<name>.<ext>`
        // pattern a plain `-l<name>` link line assumes. `dylib:+verbatim=<name>` was tried here to
        // link against the exact filename without that convention, but the `verbatim` modifier is
        // only honored by linker flavors with a literal-name `-l` syntax (GNU ld/lld's `-l:name`);
        // Apple's `ld` has none, silently falls back to its normal `-l<name>` mangling, and fails to
        // find the file. Passing the full path directly as a link arg sidesteps `-l` name-mangling
        // entirely, on every linker flavor.
        println!("cargo:rustc-link-arg={}", lib_dir.join(&dll_basename).display());
        // The path above only satisfies the *build-time* linker. NativeAOT publishes macOS dylibs
        // with their own install name set to `@rpath/<name>` (visible in `otool -D`/dyld errors),
        // so whatever path we link against, the *runtime* dependency recorded in the test binary is
        // always that `@rpath`-relative name, not our real path - `@rpath` resolves via `LC_RPATH`
        // entries in the loading binary, and without one dyld falls back to a handful of unrelated
        // default locations and never finds it. Embed lib_dir as an rpath entry so it does. Also
        // covers the equivalent (if less commonly hit) case on Linux, where `DT_NEEDED` may likewise
        // resolve to a bare name that only `DT_RUNPATH` steers back to a non-standard directory.
        println!("cargo:rustc-link-arg=-Wl,-rpath,{}", lib_dir.display());
    }
}

/// Scans `dir` for a single file with extension `ext` and returns its basename. Used for the
/// `EXCELREADER_NATIVE_LIB_DIR` override, where the exact filename of the native binary isn't
/// known ahead of time (unlike the download path, where it's always `asset_name`).
fn find_native_lib(dir: &Path, ext: &str) -> String {
    let mut candidates: Vec<String> = fs::read_dir(dir)
        .unwrap_or_else(|e| panic!("failed to read EXCELREADER_NATIVE_LIB_DIR {}: {e}", dir.display()))
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
    // A tiny hand-rolled HTTPS GET via curl/PowerShell rather than a new crate dependency
    // (reqwest, ureq, ...) - this is a one-shot build-time download, not runtime networking, and
    // every supported CI/dev platform already ships one of these tools.
    let status = if cfg!(windows) {
        Command::new("powershell")
            .args([
                "-NoProfile", "-Command",
                &format!("Invoke-WebRequest -Uri '{url}' -OutFile '{}'", dest.display()),
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

// Embedded (not read from disk at build time) so the crate published to crates.io still has it:
// `rust/excelreader/build.rs` itself gets shipped as source in the package tarball, and Cargo
// compiles it fresh at the consumer's build time - `include_str!` is resolved then too, against
// wherever this very build.rs file lives on disk. Reaching outside the crate directory (e.g.
// `../../src/ExcelReader.Native/include/excelreader-phase1.def`) does not survive that: `cargo
// package` only ships files under `rust/excelreader/`, so a path escaping the crate root doesn't
// exist in the extracted tarball and the consumer's build fails. Keeping this copy inside the
// crate directory (kept in sync with the canonical copy at
// src/ExcelReader.Native/include/excelreader-phase1.def) makes `include_str!` resolve correctly
// both in this repo checkout and in the packaged/published crate.
const PHASE1_DEF_EXPORTS: &str = include_str!("excelreader-phase1.def");

fn generate_windows_implib(
    implib: &Path,
    target_env: &str,
    arch: &str,
    out_dir: &Path,
    dll_basename: &str,
) {
    // The checked-in .def has no LIBRARY statement, so neither lib.exe nor dlltool would
    // otherwise know what DLL name to bake into the import descriptors of the generated import
    // lib - they'd fall back to the .def file's own basename (`excelreader-phase1.dll`), which is
    // wrong: the real binary is named `excelreader-native-<os>-<arch>.dll` (downloaded) or
    // whatever EXCELREADER_NATIVE_LIB_DIR points at (override, e.g. `ExcelReader.Native.dll`).
    // Generate a temporary copy of the .def with an explicit LIBRARY line naming the real file,
    // and feed that to both tools instead of the checked-in .def directly.
    let generated_def = out_dir.join("excelreader-phase1.generated.def");
    fs::write(&generated_def, format!("LIBRARY {dll_basename}\n{PHASE1_DEF_EXPORTS}"))
        .unwrap_or_else(|e| panic!("failed to write {}: {e}", generated_def.display()));

    if target_env == "msvc" {
        let lib_exe = "lib.exe";
        let status = Command::new(lib_exe)
            .arg(format!("/def:{}", generated_def.display()))
            .arg(format!("/out:{}", implib.display()))
            .arg(format!("/machine:{arch}"))
            .status()
            .unwrap_or_else(|e| panic!("failed to invoke {lib_exe} (run from a VS developer prompt): {e}"));
        assert!(status.success(), "{lib_exe} /def failed");
    } else {
        // dlltool derives names for the temporary object files it generates internally from the
        // `-l` output path, mangling path separators into underscores; with a long absolute path
        // (routine for a Cargo OUT_DIR nested several directories deep) it fails with "failed to
        // open temporary head file". Run it with `out_dir` as the working directory and pass a
        // bare relative filename for `-l` to keep that derived name short.
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
