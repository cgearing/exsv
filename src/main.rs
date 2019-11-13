extern crate csv;

use std::env;
use calamine::{Xlsx, open_workbook, Reader, Rows};
use calamine::DataType;

fn get_contents(cell: &DataType) -> String {
     match cell {
        DataType::String(cell) => { return String::from(cell) },
        DataType::Int(cell) => { return cell.to_string() },
        DataType::Float(cell) => { return cell.to_string() },
        DataType::Bool(cell) => { return cell.to_string() },
         _ => { return String::default() } ,
        };
}

fn main() {
    let args: Vec<String> = env::args().collect();
    
    let excel_file = &args[1];
    let csv_file = &args[2];
    
    let mut wtr = csv::Writer::from_path(csv_file).unwrap();
    let mut excel: Xlsx<_> = open_workbook(excel_file).unwrap();
    let sheet_names = excel.sheet_names().to_vec();
    for name in sheet_names {
        if let Some(Ok(range)) = excel.worksheet_range(&name) {
            let rows: Rows<DataType> = range.rows();
            for row in rows {
                let result = row.iter().map(|x| get_contents(x));
                Some(wtr.write_record(result));
            }  
        }
    }
}