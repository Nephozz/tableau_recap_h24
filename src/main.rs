mod read;
mod write;

use std::{io::BufReader, fs::File};
use calamine::{Reader, open_workbook, Xlsx};
use rust_xlsxwriter::Workbook;
use read::{read_sheets, get_names, get_b00_sheets};
use write::{create_dates_sheet, create_peoples_sheet, fill_dates, fill_peoples};

pub const MOIS: &str = "Août-Septembre";
pub const NB_MOIS: usize = 1;
pub const ANNEE: &str = "2023";
const FILE_PATH: &str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/09-2023-Club.xlsx";
const SAVE_PATH: &str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/";


fn main() {
    let path: &'static str = FILE_PATH;
    let mut answers: Xlsx<BufReader<File>> = open_workbook(path).expect(
        "Impossible d'ouvrir le fichier !"
    );

    let sheets_bind: Vec<String> = answers.sheet_names().to_owned();
    let sheets: Vec<&String> = sheets_bind
        .iter()
        .filter(|&s| s.contains("TVn7"))
        .collect();

    //println!("{}", sheets.len());

    let (dates_info, peoples_info, b00_info) = read_sheets(&mut answers, &sheets);
    let names_list = get_names(&peoples_info);
    let sheets_b00 = get_b00_sheets(sheets.clone(), b00_info);
    

    let mut workbook = Workbook::new();
    let workbook_name: String = "Accès ".to_owned() + MOIS + ".xlsx";

    let local = workbook.add_worksheet()
        .set_name("Local")
        .expect("Impossible de renommer la feuille \"Local\"");
    create_dates_sheet(local, &sheets);
    fill_dates(local, &sheets, &dates_info);

    let b00 = workbook.add_worksheet()
        .set_name("B00")
        .expect("Impossible de renommer la feuille \"B00\"");
    create_dates_sheet(b00, &sheets_b00);
    fill_dates(b00, &sheets_b00, &dates_info);

    let peoples = workbook.add_worksheet()
        .set_name("Personnes avec accès")
        .expect("Impossible de renommer la feuille \"Personnes avec accès\"");
    create_peoples_sheet(peoples, &sheets, &names_list);
    fill_peoples(peoples, &sheets, &peoples_info);

    //let path = std::path::Path::new(SAVE_PATH + workbook_name);
    workbook.save(SAVE_PATH.to_owned() + &workbook_name).expect("Echec de la sauvegarde !");
}