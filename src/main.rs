use calamine::{Reader, open_workbook, Xlsx};
use rust_xlsxwriter::{Format, Image, Workbook, FormatAlign, FormatBorder, Worksheet};

fn main() {
    let path: &'static str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/05-2023 (réponses).xlsx";
    let reponses: Xlsx<_> = open_workbook(path).expect("Ne peux pas ouvrir le fichier !");

    let sheets_bind: Vec<String> = reponses.sheet_names().to_owned();
    let sheets: Vec<&String> = sheets_bind
        .iter()
        .filter(|&s| s.contains("TVn7"))
        .collect();

    let mois= "Mai";

    let mut workbook = Workbook::new();

    let mut local = workbook.add_worksheet()
        .set_name("Local");

    let mut b00 = workbook.add_worksheet()
        .set_name("B00");

    let mut personnes = workbook.add_worksheet()
        .set_name("Perssonnes avec accès");

    for s in sheets {
        println!("{}", s);
    }

    workbook.save("Accès.xlsx");
}