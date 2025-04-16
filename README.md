# Excel Reader (Rust)

An Excel reader in pure Rust.

[Documentation]()

## Capabilities

Excel Reader is a pure Rust library to read and parse xlsx files.


You can use this library to get
- An overview of Sheets in the workbook
- Detail information on worksheets including dimension, merged cells, tables, and some other properties
- Cell values and styles such as border, fill, and font.


If the processed information above does not meet your needs.
You can also get the raw version (parsed xml in Rust structures) directly for the following elements.
- Stylesheet
- Workbook
- Sharedstrings
- Theme
- Worksheet
- Tables


## Installation

Run the following Cargo command in the project directory:
```
cargo add excel_reader
```


Or add the following line to `Cargo.toml`:
```
excel_reader = "0.1.0"
```


## Examples