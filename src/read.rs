use std::{io::BufReader, fs::File};
use calamine::{Reader, Xlsx, DataType, Range};

pub fn read_date(range: &Range<DataType>, i: u32, j: u32) -> Vec<String> {
        let mut dates_sheet: Vec<String> = Vec::new();
        let date = range.get_value((i,j))
            .unwrap()
            .to_owned()
            .to_string();
        let date_parse: Vec<&str> = date.split_whitespace().collect();
        dates_sheet.push(String::from(date_parse[1]));
        dates_sheet.push(String::from(date_parse[3]));
        
        //println!("{:?}", dates_sheet);
        return dates_sheet;
}

pub fn read_event() -> Vec<String> {
    let liste_event: Vec<String> = Vec::<String>::new();

    return liste_event;
}

pub fn read_peoples(range: &Range<DataType>) -> Vec<String> {
    let mut personnes_sheet : Vec<String> = Vec::new();
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
    return personnes_sheet;
}

pub fn read_sheet(workbook: &mut Xlsx<BufReader<File>>, sheets: &Vec<&String>) -> (Vec<Vec<String>>,Vec<Vec<String>>,Vec<bool>) {
    let mut info_dates: Vec<Vec<String>> = Vec::new();
    let mut info_personnes: Vec<Vec<String>> = Vec::new();
    let mut info_b00: Vec<bool> = Vec::new();

    for s in sheets {
        let range: Range<DataType> = workbook.worksheet_range(s).unwrap().unwrap();

        let mut dates_sheet: Vec<String> = read_date(&range, 1, 3);
        dates_sheet.append(&mut read_date(&range, 1, 4));

        let b00_sheet: bool;
        let b00: String = range.get_value((1,7))
            .unwrap()
            .to_owned()
            .to_string();
        if b00.contains("B00") {b00_sheet = true;}
        else {b00_sheet = false;};

        let personnes_sheet: Vec<String> = read_peoples(&range);

        info_b00.push(b00_sheet);
        info_dates.push(dates_sheet);
        info_personnes.push(personnes_sheet);
    }
    return (info_dates, info_personnes, info_b00);
}