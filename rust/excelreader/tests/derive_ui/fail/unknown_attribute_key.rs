use excelreader::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(column = "Nome")]
    nome: String,
}

fn main() {}
