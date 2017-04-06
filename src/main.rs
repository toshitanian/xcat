extern crate calamine;
extern crate csv;

#[macro_use]
extern crate clap;


use std::env;
use std::io::Write;
use std::path::PathBuf;
use calamine::{Excel, Range, DataType, Result};
use clap::{Arg, App};

macro_rules! println_stderr(
    ($($arg:tt)*) => { {
        let r = writeln!(&mut ::std::io::stderr(), $($arg)*);
        r.expect("failed printing to stderr");
    } }
);


fn read_as_excel(sce: PathBuf) {
    let mut xl = Excel::open(&sce).unwrap();

    let range_result = xl.sheet_names()
        .map(|elem| elem[0].to_string())
        .and_then(|name| xl.worksheet_range(&name));
    let range = range_result.unwrap();
    write_range(range).unwrap();
}

fn read_as_csv(sce: PathBuf, delimiter: &str) {
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


fn main() {
    let matches = App::new("xcat")
        .version(crate_version!())
        .author(crate_authors!())
        .about("Expanded cat tool especially for Excel")
        .arg(Arg::with_name("delimiter")
             .help("Delimiter for csv file")
             .short("d")
             .long("delimiter")
             .value_name("DELIMITER")
             .takes_value(true))
        .arg(Arg::with_name("input_files")
             .help("Input files to use")
             .multiple(true)
             .required(true)
             .value_name("file")
             .takes_value(true))
        .get_matches();
    let delimiter = matches.value_of("delimiter").unwrap_or(",");

    let files: Vec<_> = matches.values_of("input_files").unwrap().collect();
    for file in files {
        let sce = PathBuf::from(&file);
        if !sce.exists() {
            println_stderr!("{}: No such file or directory", file);
            continue;
        }
        match sce.extension().and_then(|s| s.to_str()) {
            Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => read_as_excel(sce),
            Some("csv") | Some("txt") => read_as_csv(sce, delimiter),
            _ => panic!("Expecting an excel file"),
        }

    }

}

fn write_range(range: Range) -> Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            let _ = match *c {
                DataType::Empty => print!(""),
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
