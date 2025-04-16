use excel_reader::{
    common_types::Coordinate, excel::Excel,
    processed::spreadsheet::sheet::worksheet::cell_value::CellValueType,
};

/// Demo for basic usages
///
/// sample Excels: https://create.microsoft.com/en-us/search?query=table&filters=excel
fn main() -> anyhow::Result<()> {
    let path = "examples/sample.xlsx";

    // excel from a reader
    // let reader = BufReader::new(File::open(path)?);
    // let mut archive = ZipArchive::new(reader)?;
    // let mut excel = Excel::from_path(path)?;

    // excel directly from path
    let mut excel = Excel::from_path(path)?;

    // get basic sheet info
    let sheets = excel.get_sheets()?;
    if sheets.is_empty() {
        println!("Excel contains no sheets");
        return Ok(());
    }
    println!("sheets: ");
    for sheet in sheets.clone() {
        println!("--------");
        println!("id: {}", sheet.sheet_id);
        println!("name: {}", sheet.name);
        println!("type: {:?}", sheet.r#type);
        println!("visiblity: {:?}", sheet.visible_state);
    }
    println!("--------");
    println!("--------");

    // get worksheet detail
    let mut worksheet = excel.get_worksheet(&sheets[2].clone())?;
    println!("worksheet: {}", worksheet.name);
    println!("dimension: {:?}", worksheet.dimension);
    println!(
        "Use 1904 backward compatibility date system: {:?}",
        worksheet.is_1904
    );
    println!(
        "Calculation reference: {:?}",
        worksheet.calculation_reference_mode
    );
    if !worksheet.clone().merged_cells.is_empty() {
        println!("merged cells: ");
        for (index, merged_cell) in worksheet.clone().merged_cells.iter().enumerate() {
            println!("{}: {:?}", index + 1, merged_cell);
        }
    }
    if !worksheet.clone().tables.is_empty() {
        println!("tables: ");
        for (index, table) in worksheet.clone().tables.iter().enumerate() {
            println!("--------");
            println!("{}: {:?}", index + 1, table.display_name);
            println!("dimension: {:?}", table.dimension);
            println!("columns: {:?}", table.columns);
            println!("table style: {:?}", table.table_style.name);
        }
        println!("--------");
    }

    // get cells (value and style) in a worksheet
    let Some(dimension) = worksheet.dimension else {
        println!("No cells available in worksheet");
        return Ok(());
    };

    let (start, end) = (dimension.start, dimension.end);
    let mut row_index = start.row;

    println!("Cells: ");

    while row_index <= end.row {
        let mut col_index = start.col;
        while col_index <= end.col {
            let cell = worksheet.get_cell(Coordinate::from_point((row_index, col_index)))?;
            println!("--------");
            println!("coordinate: {:?}", cell.coordinate);
            println!("value {:?}.", cell.value);
            if let CellValueType::Numeric(_) = cell.value {
                println!(
                    "Numeric format: {:?}",
                    cell.property.numbering_format.format_code
                )
            }
            let properties = cell.property;

            println!("size: {} * {}", properties.width, properties.height);
            println!("hidden : {:?}", properties.hidden);
            println!("show_phonetic : {:?}", properties.show_phonetic);
            println!("font : {:?}", properties.font);
            println!("border : {:?}", properties.border);
            println!("fill : {:?}", properties.fill);
            println!("alignment : {:?}", properties.alignment);
            col_index += 1;
        }
        row_index += 1;
    }

    Ok(())
}
