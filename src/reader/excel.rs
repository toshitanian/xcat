extern crate calamine;

use std::path::PathBuf;

pub fn read(sce: PathBuf, delimiter: &str) {
    let mut xl = calamine::Excel::open(&sce).unwrap();
    let range_result = xl.sheet_names()
        .map(|elem| elem[0].to_string())
        .and_then(|name| xl.worksheet_range(&name));
    let range = range_result.unwrap();
    write_range(range, delimiter);
}

fn write_range(range: calamine::Range, delimiter: &str) {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            let _ = match *c {
                calamine::DataType::Empty => print!(""),
                calamine::DataType::String(ref s) => print!("{}", s),
                calamine::DataType::Float(ref f) => print!("{}", f),
                calamine::DataType::Int(ref i) => print!("{}", i),
                calamine::DataType::Error(ref e) => print!("{:?}", e),
                calamine::DataType::Bool(ref b) => print!("{}", b),
            };
            if i != n {
                let _ = print!("{}", delimiter);
            }
        }
        let _ = println!("");
    }
}
