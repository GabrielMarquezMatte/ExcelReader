use excelreader::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row<const N: usize>
where
    String: Default,
{
    #[excel(name = "Nome")]
    nome: String,
}

fn main() {
    assert_eq!(Row::<4>::bindings()[0].names[0], "Nome");
}
