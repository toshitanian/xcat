extern crate calamine;
extern crate csv;


use std::env;
use std::path::PathBuf;
use calamine::{Excel, Range, DataType, Result};

fn read_as_excel(sce: PathBuf) {
    let mut xl = Excel::open(&sce).unwrap();

    let sheet_name = xl.sheet_names().unwrap()[0];

    let mut xl2 = Excel::open(&sce).unwrap();
    let range = xl2.worksheet_range(&sheet_name).unwrap();

    write_range(range).unwrap();
}

fn read_as_csv(sce: PathBuf) {
    let mut rdr = csv::Reader::from_file(&sce).unwrap();
    for row in rdr.records().map(|r| r.unwrap()) {
        let n = row.len() - 1;
        for (i, c) in row.iter().enumerate() {
            print!("{}", c);
            if i != n {
                let _ = print!(",");
            }

        }
        println!("");
    }
}


fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file = env::args()
        .skip(1)
        .next()
        .expect("Please provide an file to show, csv, excel");
    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => read_as_excel(sce),
        Some("csv") | Some("txt") => read_as_csv(sce),
        _ => panic!("Expecting an excel file"),
    }

}

fn write_range(range: Range) -> Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            let _ = match *c {
                DataType::Empty => println!(""),
                DataType::String(ref s) => print!("{}", s),
                DataType::Float(ref f) => print!("{}", f),
                DataType::Int(ref i) => print!("{}", i),
                DataType::Error(ref e) => print!("{:?}", e),
                DataType::Bool(ref b) => print!("{}", b),
            };
            if i != n {
                let _ = print!(",");
            }
        }
        let _ = println!("");
    }
    Ok(())
}
