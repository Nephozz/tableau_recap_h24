use rust_xlsxwriter::{Format, FormatAlign, Worksheet, XlsxColor, FormatBorder};
use crate::{MOIS, ANNEE};

pub const COLORS: [u32; 9] = [
    0x7FF584,
    0x95BDF5,
    0xF5F37D,
    0xF564A0,
    0xF5C871,
    0xB27DF5,
    0xBEFFFE,
    0xFF9F9F,
    0xFFFFFF,
];

pub fn init_jour(worksheet: &mut Worksheet) {
    for j in 1..32 {
        let format = Format::new()
            .set_align(FormatAlign::Left)
            .set_bold()
            .set_font_color(XlsxColor::White)
            .set_background_color(XlsxColor::Theme(0, 4))
            .set_border(FormatBorder::Thin);
        worksheet.write_with_format(1, j, j, &format).unwrap();
        worksheet.set_column_width(j, 3).unwrap();
    } 
}

pub fn init_noms(personnes: &mut Worksheet, liste_noms: &Vec<String>, sheets: &Vec<&String>, last_col: u16) {
    let mut j: u16 = 1;
    let mut i: u32 = 1;

    personnes.set_row_height(i, 30).unwrap();

    for nom in liste_noms {
        if (j - 1)%8 == 0 && *nom != liste_noms[0] {
            i += (sheets.len() + 1) as u32;
            personnes.set_row_height(i, 30).unwrap();
            i += 1;
            init_titre(personnes, i, last_col);
            init_table(personnes, i, last_col, sheets);
            init_event(personnes, sheets, i);
            i += 1;
            j = 1;
        }
        let format = Format::new()
            .set_align(FormatAlign::Center)
            .set_align(FormatAlign::VerticalCenter)
            .set_bold()
            .set_background_color(XlsxColor::Theme(0, 4))
            .set_font_color(XlsxColor::White)
            .set_text_wrap()
            .set_border(FormatBorder::Thin);
        personnes.write_string_with_format(i, j, &nom as &str, &format).unwrap()
            .set_column_width(j, 13).unwrap();
        j += 1;
    }
}

pub fn init_titre(worksheet: &mut Worksheet, first_row: u32, last_col: u16) {
    let format = Format::new()
        .set_background_color(XlsxColor::Theme(0, 2))
        .set_border(FormatBorder::Thin);

    worksheet.set_column_width(0, 17).unwrap()
        .merge_range(first_row, 0, first_row + 1, 0, "", &format).unwrap();
    
    let format = Format::new()
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_background_color(XlsxColor::RGB(0xCCFFCC))
        .set_border(FormatBorder::Thin);
    let titre: String = "Accès TVn7 ".to_owned() 
        + MOIS + " " 
        + ANNEE + " " 
        +"(" + &worksheet.name() + ")";
    worksheet.merge_range(first_row, 1,
        first_row, last_col,
        &titre, &format).unwrap();
}

pub fn init_date_sheet(worksheet: &mut Worksheet, sheets: &Vec<&String>) {
    let last_col: u16 = 31;

    init_event(worksheet, sheets, 0);
    init_table(worksheet, 0, last_col, sheets);
    init_jour(worksheet);
    init_titre(worksheet, 0, last_col);
}

pub fn init_personnes_sheet(worksheet: &mut Worksheet, sheets: &Vec<&String>, liste_noms: &Vec<String>) {
    let last_col: u16 = 8 as u16;

    init_event(worksheet, sheets, 0);
    init_table(worksheet, 0, last_col, sheets);
    init_noms(worksheet, liste_noms, sheets, last_col);
    init_titre(worksheet, 0, last_col);
}

pub fn init_table(worksheet: &mut Worksheet, first_row:u32, last_col: u16, sheets: &Vec<&String>) {
    let format_clair = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_background_color(XlsxColor::Theme(0, 0));
    let format_sombre = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_background_color(XlsxColor::Theme(0, 2));

    let mut i: u32 = first_row;

    for _s in sheets {
        for j in 0..last_col {
            if i%2 == 0 {
                worksheet.write_string_with_format(i + 2, j + 1, "", &format_clair).unwrap();
            } else {
                worksheet.write_string_with_format(i + 2, j + 1, "", &format_sombre).unwrap();
            }
        }
        i += 1;
    }
}

pub fn init_event(worksheet: &mut Worksheet, sheets: &Vec<&String>, first_row: u32) {
    let mut i: usize = (first_row + 2) as usize;
    let mut k = 0;

    for s in sheets {
        let slice = &s[14..s.len()];
        let format = Format::new()
            .set_background_color(XlsxColor::RGB(COLORS[k]))
            .set_border(FormatBorder::Thin);
        worksheet.write_string_with_format(i as u32, 0, slice, &format).unwrap();
        i += 1;
        k += 1;
    }
}