#![forbid(unsafe_code)]

use excelreader::{Date, ExcelMapper, Time, Timestamp};

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Nome")]
    nome: String,
    #[excel(name = "Idade")]
    idade: Option<i32>,
    #[excel(name = "Dia")]
    dia: Date,
    #[excel(name = "Hora")]
    hora: Time,
    #[excel(name = "Instante")]
    instante: Option<Timestamp>,
}

fn main() {
    assert_eq!(Row::bindings().len(), 5);
}
