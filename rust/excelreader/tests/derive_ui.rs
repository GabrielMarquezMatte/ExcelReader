#[test]
fn ui() {
    let t = trybuild::TestCases::new();
    t.pass("tests/derive_ui/pass/*.rs");
}
