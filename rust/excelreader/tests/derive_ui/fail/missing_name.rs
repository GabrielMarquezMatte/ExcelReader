use excelreader::workbook::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row {
    nome: String,
}

fn main() {}
