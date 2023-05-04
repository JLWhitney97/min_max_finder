use calamine::{open_workbook, DataType, Reader, Xlsx};
use rfd::FileDialog;
use std::str::FromStr;

fn main() {
    // Have the user select xlsx files
    let files = FileDialog::new()
        .add_filter("excel", &["xlsx"])
        .set_directory("~/Desktop")
        .pick_files()
        .unwrap();

    // Sheets will hold all of the excel files
    let mut sheets: Vec<Xlsx<_>> = Vec::new();
    // Import the files to sheets using calamine's open_workbook function.
    for f in files {
        if let Ok(s) = open_workbook(f.as_path()) {
            sheets.push(s);
        }
    }

    // Find the min and max of the "SN" column by looping over all the sheets
    let mut min_sn = i64::MAX;
    let mut max_sn = i64::MIN;
    for mut excel in sheets {
        if let Some(Ok(r)) = excel.worksheet_range("Sheet1") {
            let mut sn_index = 0;
            for (row_counter, row) in r.rows().enumerate() {
                // On the first row of the sheet, find the "SN" column header and set sn_index to that column's index
                if row_counter == 0 {
                    for (cell_counter, column_header) in row.iter().enumerate() {
                        match column_header {
                            DataType::String(val) => {
                                if val == "SN" {
                                    sn_index = cell_counter;
                                }
                            }
                            _ => continue,
                        }
                    }
                }
                // Parse out the SN value in the current row
                let value;
                match &row[sn_index] {
                    DataType::String(val) => value = i64::from_str(&val),
                    DataType::Int(val) => value = Ok(*val),
                    _ => continue,
                }
                // Min and Max comparison is done if the row was able to be parsed correctly, otherwise the cell is skipped
                match value {
                    Ok(val) => {
                        min_sn = std::cmp::min(min_sn, val);
                        max_sn = std::cmp::max(max_sn, val);
                    }
                    Err(_) => continue,
                }
            }
        }
    }
    // Output the result and keep the terminal open
    println!("Min: {}\nMax: {}", min_sn, max_sn);
    let _ = std::process::Command::new("cmd.exe")
        .arg("/c")
        .arg("pause")
        .status();
}
