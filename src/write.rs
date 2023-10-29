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


/* init_calendar : initialise le calandrier pour le récap
worksheet : &Worsheet, nouvelle feuille où l'on veut faire un récap journalier
Return : void
*/
pub fn init_calendar(worksheet: &mut Worksheet) {
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

/* init_names : initialise la liste des personnes ayant des accès pendant le mois
peoples : &mut Worksheet, nouvelle feuille où l'on veut inscrire les noms de persones ayant des accès
names_list : &Vec<String>, liste des noms (unique) des personnes ayant les accès
sheets : &Vec<&String>, ensenble des demandes triées
last_col: u16, ?
    TODO
Return : void
*/
pub fn init_names(peoples: &mut Worksheet, names_list: &Vec<String>, sheets: &Vec<&String>, last_col: u16) {
    let mut j: u16 = 1;
    let mut i: u32 = 1;

    peoples.set_row_height(i, 30).unwrap();

    for name in names_list {
        if (j - 1)%8 == 0 && *name != names_list[0] {
            i += (sheets.len() + 1) as u32;
            peoples.set_row_height(i, 30).unwrap();
            i += 1;
            write_title(peoples, i, last_col);
            init_table(peoples, i, last_col, sheets);
            init_event(peoples, sheets, i);
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
        peoples.write_string_with_format(i, j, &name as &str, &format).unwrap()
            .set_column_width(j, 13).unwrap();
        j += 1;
    }
}

/* write_title : initialise le titre de la feuille
worksheet: &mut Worksheet, feuille sur laquelle on initialise le titre
first_row: u32, 
last_col: u16,
*/
pub fn write_title(worksheet: &mut Worksheet, first_row: u32, last_col: u16) {
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
    let title: String = "Accès TVn7 ".to_owned() 
        + MOIS + " " 
        + ANNEE + " " 
        +"(" + &worksheet.name() + ")";
    worksheet.merge_range(first_row, 1,
        first_row, last_col,
        &title, &format).unwrap();
}

//TODO
pub fn create_dates_sheets(worksheet: &mut Worksheet, sheets: &Vec<&String>) {
    let last_col: u16 = 31;

    init_event(worksheet, sheets, 0);
    init_table(worksheet, 0, last_col, sheets);
    init_calendar(worksheet);
    write_title(worksheet, 0, last_col);
}

//TODO
pub fn create_peoples_sheet(worksheet: &mut Worksheet, sheets: &Vec<&String>, names_list: &Vec<String>) {
    let last_col: u16 = 8 as u16;

    init_event(worksheet, sheets, 0);
    init_table(worksheet, 0, last_col, sheets);
    init_names(worksheet, names_list, sheets, last_col);
    write_title(worksheet, 0, last_col);
}

//TODO
pub fn init_table(worksheet: &mut Worksheet, first_row:u32, last_col: u16, sheets: &Vec<&String>) {
    let bright_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_background_color(XlsxColor::Theme(0, 0));
    let dark_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_background_color(XlsxColor::Theme(0, 2));

    let mut i: u32 = first_row;

    for _s in sheets {
        for j in 0..last_col {
            if i%2 == 0 {
                worksheet.write_string_with_format(i + 2, j + 1, "", &bright_format).unwrap();
            } else {
                worksheet.write_string_with_format(i + 2, j + 1, "", &dark_format).unwrap();
            }
        }
        i += 1;
    }
}

//TODO
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

/* fill_personnes : remplit les noms des personnes ayant les accès sur le mois
worksheet : &mut Worksheet, feuille des accès par personne
sheets : ensemble des feuilles triées
peoples_info : ?
TODO
*/
pub fn fill_personnes(worksheet: &mut Worksheet, sheets: &Vec<&String>, peoples_info: &Vec<Vec<String>>) {
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
        for j in 1..peoples_info[i].len() {
            let n = j/8;
            println!("{}, {}, {}",i,j,n);
            if i%2 == 0 {_format = format_clair.clone();} 
            else {_format = format_sombre.clone();}

            worksheet.write_with_format(
                    (i + 2) as u32, 
                    get_name_col(&peoples_info[i][j], peoples_info),
                    "X", 
                    &_format)
                .unwrap();
        }
    }
}

/* fill_dates : remplit les dates dans les feuilles des accès par jours
worksheet : &mut Worksheet, feuille des accès par personne
sheets : ensemble des feuilles triées
dates_info : ?
TODO
*/
pub fn fill_dates(worksheet: &mut Worksheet, sheets: &Vec<&String>, dates_info: &Vec<Vec<String>>) {
    for i in 0..sheets.len() {
        let format = Format::new()
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin)
        .set_background_color(COLORS[i]);

        let heure_debut = dates_info[i][1].as_str();
        let heure_fin = dates_info[i][3].as_str();
        let date_debut = dates_info[i][0].parse().unwrap();
        let date_fin = dates_info[i][2].parse().unwrap();

        worksheet.write_with_format((i + 2) as u32, date_debut, heure_debut, &format).unwrap();
        worksheet.write_with_format((i + 2) as u32, date_fin, heure_fin, &format).unwrap();
        for j in (date_debut + 1)..(date_fin) {
            worksheet.write_with_format((i + 2) as u32, j, "", &format).unwrap();
        }
    }
}