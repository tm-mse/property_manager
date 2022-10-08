extern crate office;
use std::env;
use office::{Excel, DataType};
use std::path::PathBuf;

fn main() {
    
    // Take the argumnents passed by the user
    let args: Vec<String> = env::args().collect();
    let file = env::args()
        .nth(1)
        .expect("Make sure the file path is correct...");
    // Verify the extension
    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }
    
    // Read the excel file
    let mut workbook = Excel::open(&sce).unwrap();
    let sheets = workbook.sheet_names().unwrap();
    println!("Sheets: {:#?}", &sheets);
     
    for sheet in sheets {
        let range: office::Range = workbook.worksheet_range(&sheet).unwrap();
        let total_cells = range.get_size().0 * (range.get_size().1); 
        let non_empty_cells: usize = range.rows().map(|r| {r.iter().filter(|cell| cell != &&DataType::Empty).count()
            }).sum();
        println!("Found {} cells in {}, including {} non empty cells", 
                 total_cells, &sheet, non_empty_cells);
    }
    
  //  let total_cells = range.get_size().0 * range.get_size().1;
  //  let non_empty_cells: usize = range.rows().map(|r| {r.iter().filter(|cell| cell != &&DataType::Empty).count()
  //      }).sum();
  //  println!("Found {} cells in {}, including {} non empty cells", 
  //           total_cells, &sheets[0], non_empty_cells);
}

//TO DO :
//  [] Parser des tableaux exploitable
//      + nom du locataire
//      + addresse du locataire
//      + nom du propriétaire
//      + loyer payé
//      + reste du
//  [] Créer le pdf
//  [] automatiser la lecture du fichier excell et l'écriture de fichier pdf
