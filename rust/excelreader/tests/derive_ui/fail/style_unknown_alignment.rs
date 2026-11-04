use excelreader::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Valor", style(h_align = "middle"))]
    valor: f64,
}

fn main() {}
