mod read;
mod init;

use std::{io::BufReader, fs::File};
use calamine::{Reader, open_workbook, Xlsx};
use rust_xlsxwriter::{Workbook, Worksheet, Format, FormatAlign};
use read::read_sheet;
use init::{init_local_sheet, init_b00_sheet, init_personnes_sheet};

pub const MOIS: &str = "Mai";
pub const ANNEE: &str = "2023";
const FILE_PATH: &str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/05-2023 (réponses).xlsx";
const SAVE_PATH: &str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/";

pub fn get_names(info_personnes: &Vec<Vec<String>>) -> Vec<String> {
    let mut liste_complete: Vec<String> = info_personnes.clone()
        .into_iter()
        .flatten()
        .collect();
    liste_complete.sort();
    liste_complete.dedup();

    return liste_complete;
}

pub fn get_name_col(name: &String, info_personnes: &Vec<Vec<String>>) -> u16 {
    let mut col: u16 = 0;
    let reference = get_names(info_personnes);
    for i in 0..reference.len() {
        if name == &reference[i] {
            col = i as u16;
            break;
        }
    }
    return col + 1;
}

pub fn fill_personnes(worksheet: &mut Worksheet, sheets: &Vec<&String>, info_personnes: &Vec<Vec<String>>) {
    let format = Format::new().set_align(FormatAlign::CenterAcross);
    for i in 0..sheets.len() {
        for j in 0..info_personnes[i].len() {
            worksheet.write_with_format((i + 2) as u32, get_name_col(&info_personnes[i][j], info_personnes), "X", &format).unwrap();
        }
    }
}

fn main() {
    let path: &'static str = FILE_PATH;
    let mut reponses: Xlsx<BufReader<File>> = open_workbook(path).expect(
        "Impossible d'ouvrir le fichier !"
    );

    let sheets_bind: Vec<String> = reponses.sheet_names().to_owned();
    let sheets: Vec<&String> = sheets_bind
        .iter()
        .filter(|&s| s.contains("TVn7"))
        .collect();

    let (info_dates, info_personnes, info_b00) = read_sheet(&mut reponses, &sheets);
    let liste_noms = get_names(&info_personnes);
    drop(info_dates);
    println!("{:#?}", info_personnes);

    let mut workbook = Workbook::new();
    let workbook_name: String = "Accès ".to_owned() + MOIS + ".xlsx";

    let local = workbook.add_worksheet()
        .set_name("Local")
        .expect("Impossible de renommer la feuille \"Local\"");
    init_local_sheet(local, &sheets, MOIS, ANNEE);

    let b00 = workbook.add_worksheet()
        .set_name("B00")
        .expect("Impossible de renommer la feuille \"B00\"");
    init_b00_sheet(b00, &sheets, MOIS, ANNEE, &info_b00);

    let personnes = workbook.add_worksheet()
        .set_name("Personnes avec accès")
        .expect("Impossible de renommer la feuille \"Personnes avec accès\"");
    init_personnes_sheet(personnes, &sheets, MOIS, ANNEE, &liste_noms);
    fill_personnes(personnes, &sheets, &info_personnes);

    workbook.save(SAVE_PATH.to_owned() + &workbook_name).expect("Echec de la sauvegarde !");
}