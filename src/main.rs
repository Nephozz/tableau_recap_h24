use std::{io::BufReader, fs::File};
use calamine::{Reader, open_workbook, Xlsx, DataType, Range};
use rust_xlsxwriter::{Format, Workbook, FormatAlign, Worksheet, XlsxColor, FormatBorder};

const MOIS: &str = "Mai";
const ANNEE: &str = "2023";
const FILE_PATH: &str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/05-2023 (réponses).xlsx";
const SAVE_PATH: &str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/";

#[derive(Debug)]
enum Info {
    Date(Vec<String>),
    B00(bool),
    Personne(Vec<String>),
}

pub fn read_date(range: &Range<DataType>, i: u32, j: u32, dates_sheet: &mut Vec<String>) {
        let date = range.get_value((i,j))
            .unwrap()
            .to_owned()
            .to_string();
        let date_parse: Vec<&str> = date.split_whitespace().collect();
        dates_sheet.push(String::from(date_parse[1]));
        dates_sheet.push(String::from(date_parse[3]));
}

pub fn read_peoples(range: &Range<DataType>, personnes_sheet: &mut Vec<String>) {
    for i in 4..11 {
        for j in 7..9 {
            let personnes_bind = range.get_value((i,j));
            match personnes_bind {
                None | Some(DataType::Empty) => {},
                Some(value) => {
                    let value = Some(value)
                        .unwrap()
                        .to_owned()
                        .to_string();
                    personnes_sheet.push(value);
                },
            };
        }
    }
}

pub fn read_sheet(workbook: &mut Xlsx<BufReader<File>>, sheets: &Vec<&String>) {
    let mut info: Vec<Info> = Vec::new();

    for s in sheets {
        let b00_sheet: bool;
        let mut dates_sheet: Vec<String> = Vec::new();
        let mut personnes_sheet: Vec<String> = Vec::new();

        let range: Range<DataType> = workbook.worksheet_range(s).unwrap().unwrap();

        read_date(&range, 1, 3, &mut dates_sheet);
        read_date(&range, 1, 4, &mut dates_sheet);

        let b00: String = range.get_value((1,7))
            .unwrap()
            .to_owned()
            .to_string();
        if b00.contains("B00") {
            b00_sheet = true;
        } else { b00_sheet = false; };

        read_peoples(&range, &mut personnes_sheet);

        info.push(Info::B00(b00_sheet));
        info.push(Info::Date(dates_sheet));
        info.push(Info::Personne(personnes_sheet));
    }
    println!("{:#?}", info[5]);
}
    
pub fn init_sheet(worksheet: &mut Worksheet, sheets: &Vec<&String>, mois: &str, annee: &str) {
        let colors: Vec<u32> = vec![
            0x7FF584,
            0x95BDF5,
            0xF5F37D,
            0xF564A0,
            0xF5C871,
            0xB27DF5,
        ];
    
        let mut i: usize = 2;
        worksheet.set_column_width(0, 17).unwrap();

        for s in sheets {
            let slice = &s[14..s.len()];
            let format = Format::new()
                .set_background_color(XlsxColor::RGB(colors[i - 2]))
                .set_border(FormatBorder::Thin);
            worksheet.write_string_with_format(i as u32, 0, slice, &format).unwrap();
            i += 1;
        }

        for j in 1..32 {
            let format = Format::new()
                .set_align(FormatAlign::Left)
                .set_bold()
                .set_font_color(XlsxColor::White)
                .set_background_color(XlsxColor::Gray)
                .set_border(FormatBorder::Thin);
            worksheet.write_with_format(1, j, j, &format).unwrap();
            worksheet.set_column_width(j, 5).unwrap();
        } 

        let format = Format::new()
            .set_bold()
            .set_font_color(XlsxColor::White)
            .set_align(FormatAlign::Center)
            .set_background_color(XlsxColor::Silver)
            .set_border(FormatBorder::Thin);
        let titre: String = "Accès TVn7 ".to_owned() 
            + mois + " " 
            + annee + " " 
            +"(" + &worksheet.name() + ")";
        worksheet.merge_range(0, 1,
            0, 30,
            &titre, &format).unwrap();
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

    read_sheet(&mut reponses, &sheets);

    let mut workbook = Workbook::new();
    let workbook_name: String = "Accès ".to_owned() + MOIS + ".xlsx";

    let local = workbook.add_worksheet()
        .set_name("Local")
        .expect("Impossible de renommer la feuille \"Local\"");
    init_sheet(local, &sheets, MOIS, ANNEE);

    let b00 = workbook.add_worksheet()
        .set_name("B00")
        .expect("Impossible de renommer la feuille \"B00\"");
    init_sheet(b00, &sheets, MOIS, ANNEE);

    let personnes = workbook.add_worksheet()
        .set_name("Perssonnes avec accès")
        .expect("Impossible de renommer la feuille \"Personnes avec accès\"");
    init_sheet(personnes, &sheets, MOIS, ANNEE);

    workbook.save(SAVE_PATH.to_owned() + &workbook_name).expect("Echec de la sauvegarde !");
}