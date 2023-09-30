mod read;
mod init;

use std::{io::BufReader, fs::File};
use calamine::{Reader, open_workbook, Xlsx};
use rust_xlsxwriter::{Workbook, Worksheet, Format, FormatAlign, FormatBorder, XlsxColor};
use read::read_sheet;
use init::{init_date_sheet, init_personnes_sheet, COLORS};

pub const MOIS: &str = "Août-Septembre";
pub const NB_MOIS: usize = 1;
pub const ANNEE: &str = "2023";
const FILE_PATH: &str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/09-2023-Club.xlsx";
const SAVE_PATH: &str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/";

pub fn get_b00_sheets(sheets: Vec<&String>, info_b00: Vec<bool>) -> Vec<&String> {
    let mut sheets_b00 = Vec::new();
    
    for i in 0..sheets.len() {
        if info_b00[i] {
            sheets_b00.push(sheets[i])
        }
    }
    return sheets_b00;
}

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
    let mut _format = Format::new();

    let format_clair = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_background_color(XlsxColor::Theme(0, 0));
    let format_sombre = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_background_color(XlsxColor::Theme(0, 2));

    for i in 0..sheets.len() {
        for j in 1..info_personnes[i].len() {
            let n = j/8;
            println!("{}, {}, {}",i,j,n);
            if i%2 == 0 {_format = format_clair.clone();} 
            else {_format = format_sombre.clone();}

            worksheet.write_with_format(
                    (i + 2) as u32, 
                    get_name_col(&info_personnes[i][j], info_personnes),
                    "X", 
                    &_format)
                .unwrap();
        }
    }
}

pub fn fill_dates(worksheet: &mut Worksheet, sheets: &Vec<&String>, info_dates: &Vec<Vec<String>>) {
    for i in 0..sheets.len() {
        let format = Format::new()
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin)
        .set_background_color(COLORS[i]);

        let heure_debut = info_dates[i][1].as_str();
        let heure_fin = info_dates[i][3].as_str();
        let date_debut = info_dates[i][0].parse().unwrap();
        let date_fin = info_dates[i][2].parse().unwrap();

        worksheet.write_with_format((i + 2) as u32, date_debut, heure_debut, &format).unwrap();
        worksheet.write_with_format((i + 2) as u32, date_fin, heure_fin, &format).unwrap();
        for j in (date_debut + 1)..(date_fin) {
            worksheet.write_with_format((i + 2) as u32, j, "", &format).unwrap();
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

    //println!("{}", sheets.len());

    let (info_dates, info_personnes, info_b00) = read_sheet(&mut reponses, &sheets);
    let liste_noms = get_names(&info_personnes);
    let sheets_b00 = get_b00_sheets(sheets.clone(), info_b00);
    

    let mut workbook = Workbook::new();
    let workbook_name: String = "Accès ".to_owned() + MOIS + ".xlsx";

    let local = workbook.add_worksheet()
        .set_name("Local")
        .expect("Impossible de renommer la feuille \"Local\"");
    init_date_sheet(local, &sheets);
    fill_dates(local, &sheets, &info_dates);

    let b00 = workbook.add_worksheet()
        .set_name("B00")
        .expect("Impossible de renommer la feuille \"B00\"");
    init_date_sheet(b00, &sheets_b00);
    fill_dates(b00, &sheets_b00, &info_dates);

    let personnes = workbook.add_worksheet()
        .set_name("Personnes avec accès")
        .expect("Impossible de renommer la feuille \"Personnes avec accès\"");
    init_personnes_sheet(personnes, &sheets, &liste_noms);
    fill_personnes(personnes, &sheets, &info_personnes);

    //let path = std::path::Path::new(SAVE_PATH + workbook_name);
    workbook.save(SAVE_PATH.to_owned() + &workbook_name).expect("Echec de la sauvegarde !");
}