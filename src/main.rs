#[macro_use]
extern crate clap;

mod reader;
use reader::excel;
use reader::csv;

use std::io::Write;
use std::path::PathBuf;
use clap::{Arg, App};

macro_rules! println_stderr(
    ($($arg:tt)*) => { {
        let r = writeln!(&mut ::std::io::stderr(), $($arg)*);
        r.expect("failed printing to stderr");
    } }
);


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
            Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => excel::read(sce, delimiter),
            Some("csv") | Some("txt") => csv::read(sce, delimiter),
            Some("ods") => println_stderr!("{}: .ods is not supported yet", file),
            _ => println_stderr!("{}: Not supported file format", file),
        }
    }
}
