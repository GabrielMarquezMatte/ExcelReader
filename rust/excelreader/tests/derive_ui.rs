//! Compile-time behavior of `#[derive(ExcelMapper)]`.
//!
//! `pass/` must compile and run; `fail/` must be rejected with the exact diagnostic recorded in the
//! sibling `.stderr` file. Only macro-emitted (`syn::Error`) diagnostics belong in `fail/` - rustc's
//! own type errors get reworded between releases, which would make this suite fail on a toolchain
//! bump rather than on a real regression. Regenerate the expectations with
//! `TRYBUILD=overwrite cargo test --test derive_ui`.

#[test]
fn ui() {
    let t = trybuild::TestCases::new();
    t.pass("tests/derive_ui/pass/*.rs");
    t.compile_fail("tests/derive_ui/fail/*.rs");
}
