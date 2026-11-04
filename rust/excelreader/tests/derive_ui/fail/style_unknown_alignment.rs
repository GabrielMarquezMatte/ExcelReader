use excelreader::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Valor", style(horizontal_alignment = "middle"))]
    valor: f64,
}

fn main() {}
