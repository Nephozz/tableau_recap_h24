use rust_xlsxwriter::{Format, FormatAlign, Worksheet, XlsxColor, FormatBorder};

pub const COLORS: [u32; 6] = [
    0x7FF584,
    0x95BDF5,
    0xF5F37D,
    0xF564A0,
    0xF5C871,
    0xB27DF5,
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
        worksheet.set_column_width(j, 5).unwrap();
    } 
}

pub fn init_noms(personnes: &mut Worksheet, liste_noms: &Vec<String>) {
let mut j: u16 = 1;
    for nom in liste_noms {
        let format = Format::new()
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter)
                .set_bold()
                .set_background_color(XlsxColor::Theme(0, 4))
                .set_font_color(XlsxColor::White)
                .set_text_wrap()
                .set_border(FormatBorder::Thin);
            personnes.write_string_with_format(1, j, &nom as &str, &format).unwrap()
                .set_column_width(j, 12).unwrap()
                .set_row_height(1, 30).unwrap();
            j += 1;
    }
}

pub fn init_titre(worksheet: &mut Worksheet, last_col: u16, mois: &str, annee: &str) {
    let format = Format::new().set_background_color(XlsxColor::Theme(0, 2));
    worksheet.set_column_width(0, 17).unwrap()
        .merge_range(0, 0, 1, 0, "", &format).unwrap();
    
    let format = Format::new()
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_background_color(XlsxColor::RGB(0xCCFFCC))
        .set_border(FormatBorder::Thin);
    let titre: String = "Accès TVn7 ".to_owned() 
        + mois + " " 
        + annee + " " 
        +"(" + &worksheet.name() + ")";
    worksheet.merge_range(0, 1,
        0, last_col,
        &titre, &format).unwrap();
}

pub fn init_local_sheet(worksheet: &mut Worksheet, sheets: &Vec<&String>, mois: &str, annee: &str) {
    let last_col: u16 = 31;

    let mut i: usize = 2;
    for s in sheets {
        let slice = &s[14..s.len()];
        let format = Format::new()
            .set_background_color(XlsxColor::RGB(COLORS[i - 2]))
            .set_border(FormatBorder::Thin);
        worksheet.write_string_with_format(i as u32, 0, slice, &format).unwrap();
        i += 1;
    }

    init_jour(worksheet);
    init_titre(worksheet, last_col, mois, annee);
}

pub fn init_b00_sheet(worksheet: &mut Worksheet, sheets: &Vec<&String>, mois: &str, annee: &str, info_b00: &Vec<bool>) {
    let last_col: u16 = 31;

    let mut i: usize = 2;
    let mut k: usize = 0;
    for s in sheets {
        if info_b00[k] {
            let slice = &s[14..s.len()];
            let format = Format::new()
                .set_background_color(XlsxColor::RGB(COLORS[k]))
                .set_border(FormatBorder::Thin);
            worksheet.write_string_with_format(i as u32, 0, slice, &format).unwrap();
            i += 1;
        }
        k += 1;
    }

    init_jour(worksheet);
    init_titre(worksheet, last_col, mois, annee);
}

pub fn init_personnes_sheet(worksheet: &mut Worksheet, sheets: &Vec<&String>, mois: &str, annee: &str, liste_noms: &Vec<String>) {
    let last_col: u16 = liste_noms.len() as u16;

    let mut i: usize = 2;
    for s in sheets {
        let slice = &s[14..s.len()];
        let format = Format::new()
            .set_background_color(XlsxColor::RGB(COLORS[i - 2]))
            .set_border(FormatBorder::Thin);
        worksheet.write_string_with_format(i as u32, 0, slice, &format).unwrap();
        i += 1;
    }

    init_noms(worksheet, liste_noms);
    init_titre(worksheet, last_col, mois, annee);
}