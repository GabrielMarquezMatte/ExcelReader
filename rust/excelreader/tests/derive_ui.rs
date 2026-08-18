#[test]
fn ui() {
    let t = trybuild::TestCases::new();
    t.pass("tests/derive_ui/pass/*.rs");
    t.compile_fail("tests/derive_ui/fail/*.rs");
}
