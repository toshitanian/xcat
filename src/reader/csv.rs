extern crate csv;

use std::path::PathBuf;

pub fn read(sce: PathBuf, delimiter: &str) {
    let mut rdr = csv::Reader::from_file(&sce).unwrap();
    for row in rdr.records().map(|r| r.unwrap()) {
        let n = row.len() - 1;
        for (i, c) in row.iter().enumerate() {
            print!("{}", c);
            if i != n {
                let _ = print!("{}", delimiter);
            }

        }
        println!("");
    }
}
