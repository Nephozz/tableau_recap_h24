use calamine::{Reader, open_workbook, Xlsx, DataType};

fn main() {
let path: &'static str = "C:/Users/thoma/OneDrive/Documents/Internet/H24/05-2023 (réponses).xlsx";
let mut workbook: Xlsx<_> = open_workbook(path).expect("Ne peux pas ouvrir le fichier !");

let sheets = workbook.sheet_names();
let mut i: i32 = 0;

for s in sheets {
    if s.contains("TVn7") {
        i += 1
    }
}

println!("{}",i);
}