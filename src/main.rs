use calamine::{Reader, open_workbook, Xlsx};
use rust_xlsxwriter::{Format, Workbook, FormatAlign, Worksheet, XlsxColor, FormatBorder};

pub fn init_table(
    worksheet: &mut Worksheet,
    sheets: &Vec<&String>,
    mois: &str,
    annee: &str)
    {
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
    let path: &'static str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/05-2023 (réponses).xlsx";
    let reponses: Xlsx<_> = open_workbook(path).expect(
        "Impossible d'ouvrir le fichier !"
    );

    let sheets_bind: Vec<String> = reponses.sheet_names().to_owned();
    let sheets: Vec<&String> = sheets_bind
        .iter()
        .filter(|&s| s.contains("TVn7"))
        .collect();

    let mois: &str = "Mai";
    let annee: &str = "2023";

    let mut workbook = Workbook::new();

    let local = workbook.add_worksheet()
        .set_name("Local").unwrap();
    init_table(local, &sheets, mois, annee);

    let b00 = workbook.add_worksheet()
        .set_name("B00").unwrap();
    init_table(b00, &sheets, mois, annee);

    let personnes = workbook.add_worksheet()
        .set_name("Perssonnes avec accès").unwrap();
    init_table(personnes, &sheets, mois, annee);

    workbook.save(
        "Accès ".to_owned() + mois + ".xlsx"
    ).unwrap();
}