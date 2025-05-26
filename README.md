# Excel Reader (Rust)

An Excel reader in pure Rust.

[Crate.io](https://crates.io/crates/excel_reader)

## Capabilities

Excel Reader is a pure Rust library to read and parse xlsx files.


You can use this library to get
- An overview of Sheets in the workbook
- Detail information on worksheets including dimension, merged cells, tables, and some other properties
- Cell values, formatting, and styles including hyperlink, border, fill, font, alignment and etc
- Worksheet drawings (Shape, Image, Picture, GraphicFrame, and GroupShape), their visual properties (position, size, geometry, fills, outlines, effects, and etc.) and non-visaul properties (locks, macros, hyperlinks, and etc.).


If the processed information above does not meet your needs, you can also get the raw version (parsed xml in Rust structures) directly for the following elements.
- Workbook and Worksheet Relationships
- Stylesheet
- Workbook
- Sharedstrings
- Theme
- Worksheet
- Tables
- Drawings


## Installation

Run the following Cargo command in the project directory:
```
cargo add excel_reader
```


Or add the following line to `Cargo.toml`:
```
excel_reader = "0.1.4"
```

## Features

### Serde
Serialization and Deserialization on processed structs can be enabled by adding the `serde` feature.
```
excel_reader = { version = "2.0.0", features = ["serde"] }
```

### Drawing
Ability on obtaining worksheet drawings can be enable by addding the `drawing` feature.
```
excel_reader = { version = "2.0.0", features = ["drawing"] }
```



## Examples

### Basic Usage
#### Initialization
To create an `Excel` structure, a representation of the zipped excel file and is what use to retireve further information, we can either provide a path, or a reader that implements `Read` and `Seek`.
```
let path = "examples/sample.xlsx";

// excel from a reader
let reader = BufReader::new(File::open(path)?);
let mut archive = ZipArchive::new(reader)?;
let mut excel = Excel::from_path(path)?;

// excel directly from path
let mut excel = Excel::from_path(path)?;
```

#### Usage
Here is how we can get sheets within the workbook, worksheet details, and cell information (value, format, and styles).

```
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

println!("Cells: ");

let cells = worksheet.get_cells()?;
for cell in cells {
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
    println!("hyperlink : {:?}", properties.hyperlink);
    println!("font : {:?}", properties.font);
    println!("border : {:?}", properties.border);
    println!("fill : {:?}", properties.fill);
    println!("alignment : {:?}", properties.alignment);
}

```


### Getting Worksheet drawings

```
let drawings = worksheet.get_drawings();
println!("drawings: {}", drawings.len());

for (_, drawing) in drawings.into_iter().enumerate() {
    println!("---------------");
    println!(
        "anchor: {}",
        serde_json::to_string_pretty(&drawing.anchor.clone())?
    );
    println!(
        "content: {}",
        serde_json::to_string_pretty(&drawing.content.clone())?
    );
}
```


### Getting Raw (Parsed XML)
If you want to write the processing logic to determine the style/format/value by yourself, there is also a list of functions provided to get the raw structures.

* no additional processing is done, XML to Rust Structure and that's it!


```
use excel_reader::excel::Excel;

/// Demo for getting raw rust structure directly parsed from XML
///
/// No additional processing done in addition to parsing.
fn main() -> anyhow::Result<()> {
    let path = "examples/sample.xlsx";
    let mut excel = Excel::from_path(path)?;

    // Get stylesheet parsed from xl/styles.xml
    let _stylesheet = excel.get_raw_stylesheet()?;

    // Get theme used parsed from get stylesheet parsed from xl/theme/theme{}.xml
    let _theme = excel.get_raw_theme()?;

    // Get shared string parsed from xl/sharedStrings.xml
    let _shared_strings = excel.get_raw_shared_strings()?;

    // Get workbook relationships
    let _workbook_relationships = excel.get_raw_workbook_relationship();

    // Get workbook parsed from xl/workbook.xml
    let _workbook = excel.get_raw_workbook()?;

    // Get a specific worksheet parsed from xl/worksheets/sheet{}.xml
    // `get_raw_worksheet_with_sheet_id` function or `get_raw_worksheet` function is also available.
    let _worksheet = excel.get_raw_worksheet_with_sheet_id(&1)?;

    // Get all tables defined in a worksheet parsed from xl/tables/table{}.xml, ..., xl/tables/table{n}.xml
    // `get_raw_tables_for_worksheet_with_name` function or `get_raw_tables_for_worksheet` function is also available
    let _tables = excel.get_raw_tables_for_worksheet_with_sheet_id(&1)?;

    // Get sheet relationships
    // `get_raw_sheet_relationship_with_name` function or `get_raw_sheet_relationship` function is also available
    // NOTE: if the sheet does not link with any target package or external resource, xl/worksheets/_rels/sheet{}.xml.rels might not exist.
    let _sheet_relationships = excel.get_raw_sheet_relationship_with_sheet_id(&1)?;

    Ok(())
}
```