use std::{io::BufReader, fs::File};
use calamine::{Reader, Xlsx, DataType, Range};

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
    println!("{:?}", personnes_sheet);
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
}