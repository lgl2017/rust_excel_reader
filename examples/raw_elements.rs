use excel_reader::excel::Excel;

/// Demo for getting raw rust structure directly parsed from XML
///
/// No additional processing done in addition to parsing.
fn main() -> anyhow::Result<()> {
    let path = "examples/table.xlsx";
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
    let _worksheet = excel.get_raw_worksheet_with_sheet_id(&6)?;

    // Get all tables defined in a worksheet parsed from xl/tables/table{}.xml, ..., xl/tables/table{n}.xml
    // `get_raw_tables_for_worksheet_with_name` function or `get_raw_tables_for_worksheet` function is also available
    let _tables = excel.get_raw_tables_for_worksheet_with_sheet_id(&6)?;

    // Get sheet relationships
    // `get_raw_sheet_relationship_with_name` function or `get_raw_sheet_relationship` function is also available
    // NOTE: if the sheet does not link with any target package or external resource, xl/worksheets/_rels/sheet{}.xml.rels might not exist.
    let _sheet_relationships = excel.get_raw_sheet_relationship_with_sheet_id(&6)?;

    Ok(())
}
