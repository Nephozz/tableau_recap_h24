use std::{io::BufReader, fs::File};
use calamine::{Reader, Xlsx, DataType, Range};


/* read_date : lit un date et l'heure de début/fin d'une demande
range : &Range<DateType>, 
TODO
i : u32, numéro de ligne de la cellule lue
j : u32, numéro de colonne de la cellule lue
Return : Vec<String>, Liste contenant une date et une heure
*/
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

/* read_events : ?
TODO
 */
pub fn _read_events() -> Vec<String> {
    let events_list: Vec<String> = Vec::<String>::new();

    return events_list;
}

/* read_state : donne l'état d'une demande (acceptée, refusée ou acceptée sous condtions)
    acceptée = true
    refusée = false
    acceptée = true avec conditions en message
range : &Range<DataType,
TODO
Return : un booléen correspondant
 */
pub fn _read_state(_range: &Range<DataType>) -> bool {
    //TODO
    return true;
}


/* read_peoples : lit les noms des personnes ayant les accès pour une demande
range : &Range<DataType>,
TODO
Return : Vec<String>, liste des personnes ayant les accès
*/
pub fn read_peoples(range: &Range<DataType>) -> Vec<String> {
    let mut peoples_sheet : Vec<String> = Vec::new();
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
                    peoples_sheet.push(value);
                },
            };
        }
    }
    return peoples_sheet;
}

/* read_sheets : lit l'ensemble des feuilles passée en argument afin d'en connaitre les dates d'une demande, les personnes ayant les accès 
ainsi que si la demande contient le B00
workbook : &Xlsx<BufReader<File>>, tableur excel lu
sheets : &Vec<&String>, ensenble des demandes triées
Return : dates_info : Vec<Vec<String>>, tableaux contenant les dates de début et fin de chaque demande
    peoples_info : Vec<Vec<String>>, tableaux contenant les personnes ayant les accès pour chaque demande
    b00_info : Vec<bool>, liste de booléen représentant si le B00 est inclus dans la demande
 */
pub fn read_sheets(workbook: &mut Xlsx<BufReader<File>>, sheets: &Vec<&String>) -> (Vec<Vec<String>>,Vec<Vec<String>>,Vec<bool>) {
    let mut dates_info: Vec<Vec<String>> = Vec::new();
    let mut peoples_info: Vec<Vec<String>> = Vec::new();
    let mut b00_info: Vec<bool> = Vec::new();

    for s in sheets {
        let range: Range<DataType> = workbook.worksheet_range(s).unwrap().unwrap();

        let mut dates_sheet: Vec<String> = read_date(&range, 1, 3);
        dates_sheet.append(&mut read_date(&range, 1, 4));

        let is_b00_sheet: bool;
        let b00: String = range.get_value((1,7))
            .unwrap()
            .to_owned()
            .to_string();
        if b00.contains("B00") {is_b00_sheet = true;}
        else {is_b00_sheet = false;};

        let peoples_sheet: Vec<String> = read_peoples(&range);

        b00_info.push(is_b00_sheet);
        dates_info.push(dates_sheet);
        peoples_info.push(peoples_sheet);
    }
    return (dates_info, peoples_info, b00_info);
}

/* get_b00_sheets : récupère les feuilles qui contienent une demande nécessitant le b00
sheets : &Vec<&String>, ensemble des demandes triées
b00_info : Vec<bool>, liste de booléen représentant si le B00 est inclus dans la demande
Return :  sheets_b00 : Vec<&String>, ensemble des feuilles contenat le b00 en demande
*/
pub fn get_b00_sheets(sheets: Vec<&String>, b00_info: Vec<bool>) -> Vec<&String> {
    let mut sheets_b00 = Vec::new();
    
    for i in 0..sheets.len() {
        if b00_info[i] {
            sheets_b00.push(sheets[i])
        }
    }
    return sheets_b00;
}

/* get_names : récupère, trie et élimine les doublons dan une liste de noms
peoples_info : ?
TODO 
Return : Vec<String>, liste trée et sans doublons de noms
*/
pub fn get_names(peoples_info: &Vec<Vec<String>>) -> Vec<String> {
    let mut full_list: Vec<String> = peoples_info.clone()
        .into_iter()
        .flatten()
        .collect();
    full_list.sort();
    full_list.dedup();

    return full_list;
}

/* get_name_col : ?
TODO
*/
pub fn get_name_col(name: &String, peoples_info: &Vec<Vec<String>>) -> u16 {
    let mut col: u16 = 0;
    let reference = get_names(peoples_info);
    for i in 0..reference.len() {
        if name == &reference[i] {
            col = i as u16;
            break;
        }
    }
    return col + 1;
}