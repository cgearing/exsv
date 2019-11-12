extern crate csv;

use calamine::{Xlsx, open_workbook, Reader, Rows};
use calamine::DataType;

fn main() {
    let mut excel: Xlsx<_> = open_workbook("Book1.xlsx").unwrap();
    let sheet_names = excel.sheet_names().to_vec();
    for name in sheet_names {
        if let Some(Ok(range)) = excel.worksheet_range(&name) {
            let rows: Rows<DataType> = range.rows();
            for row in rows {
                println!("{:?}", row);
            }
        }
    }
}

