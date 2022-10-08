extern crate office;
use std::env;
use std::path::PathBuf;
use calamine::{open_workbook_auto, Reader, Range, DataType};
use std::io::{BufWriter, Write};
use std::fs::File;
use polars::prelude::*;

fn main() { 
    
    // Take the argumnents passed by the user
    let _args: Vec<String> = env::args().collect();
    let file = env::args()
        .nth(1)
        .expect("Make sure the file path is correct...");
    // Verify the extension
    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }
    
    let mut xl = open_workbook_auto(&sce).unwrap(); 
    let iter = xl.worksheets();
    for t in iter {
        let file_path = PathBuf::from(&t.0);
        let _dest = file_path.with_extension("csv");
        let mut dest = BufWriter::new(File::create(&_dest).unwrap());

        write_range(&mut dest, &t.1).unwrap_or_else(|error| {
                print!("Cannot retrieve csv file : {}", t.0);
            });

        let df = get_df(_dest);
        println!("{:?}", df);

    }
}

fn write_range<W: Write>(dest: &mut W, range: &Range<DataType>) -> std::io::Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match *c {
                DataType::Empty => Ok(()),
                DataType::String(ref s) => write!(dest, "{}", s),
                DataType::Float(ref f) | DataType::DateTime(ref f) => write!(dest, "{}", f),
                DataType::Int(ref i) => write!(dest, "{}", i),
                DataType::Error(ref e) => write!(dest, "{:?}", e),
                DataType::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, ";")?;
            }
        }
        write!(dest, "\r\n")?;
    }
    Ok(())
}

fn get_df(path: PathBuf) -> PolarsResult<DataFrame> {
    CsvReader::from_path(path)?
        .has_header(true)
        .finish()
}
//TO DO :
//  [X]transformer tous les sheets en fichier csv
//  [] Parser des tableaux exploitable
//      + nom du locataire
//      + addresse du locataire
//      + nom du propriétaire
//      + loyer payé
//      + reste du
//  [] Créer le pdf
//  [] automatiser la lecture du fichier excell et l'écriture de fichier pdf
