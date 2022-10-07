use std::env;
use office::{Excel, Range, DataType};
use std::path::PathBuf;
use std::collections::HashMap;

fn main() {
    
    let args: Vec<String> = env::args().collect();
    let file = env::args()
        .nth(1)
        .expect("Make sure the file path is correct...");

    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let mut workbook = Excel::open(&sce).unwrap();
    let sheets = workbook.sheet_names().unwrap();
    println!("Sheets: {:#?}", &sheets);
   
    let mut ranges_sheets : HashMap::<&Range, String>  = HashMap::new();
    for sheet in sheets {
        ranges_sheets.insert(&workbook.worksheet_range(&sheet).unwrap_or_else(|error| {
                panic!("{}",format!("{}", error.kind()));
            }), sheet);
    }
   // let range = workbook.worksheet_range(&sheets[0]).unwrap_or_else(|error| {
   //             panic!("{}",format!("{}", error.kind()));
   //         });
    
    for (range, sheet) in ranges_sheets {
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
