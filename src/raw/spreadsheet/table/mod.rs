use std::io::{Read, Seek};

use anyhow::bail;
use quick_xml::events::Event;
use table_column::{load_table_columns, TableColumns};
use table_style_info::TableStyleInfo;
use zip::ZipArchive;

use crate::{
    common_types::{Coordinate, Dimension},
    excel::xml_reader,
    helper::{a1_dimension_to_row_col, string_to_bool, string_to_unsignedint},
};

use super::filter::{auto_filter::AutoFilter, sort_state::SortState};

pub mod calculated_column_formula;
pub mod table_column;
pub mod table_style_info;
pub mod totals_row_formula;
pub mod xml_column_properties;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.table?view=openxml-3.0.1
///
/// This element is the root element for a table that is not a single cell XML table.
///
/// Example:
/// ```
/// <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="financials"
/// displayName="financials" ref="A1:P701" totalsRowShown="0" headerRowDxfId="14" dataDxfId="13"
/// headerRowCellStyle="Currency" dataCellStyle="Currency">
///     <autoFilter ref="A1:P701" />
///     <tableColumns count="16">
///         <tableColumn id="1" name="Segment" />
///         <tableColumn id="2" name="Country" />
///         <tableColumn id="16" name="Product" dataDxfId="12" dataCellStyle="Currency" />
///     </tableColumns>
///     <tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" />
/// </table>
/// ```
///
/// table (Table)
#[derive(Debug, Clone, PartialEq)]
pub struct Table {
    /// extLst (Future Feature Data Storage Area)	Not supported

    // Child Elements
    /// autoFilter (AutoFilter Settings)
    pub auto_filter: Option<AutoFilter>,

    /// sortState (Sort State)
    pub sort_state: Option<SortState>,

    /// tableColumns (Table Columns)
    pub table_columns: Option<TableColumns>,

    /// tableStyleInfo (Table Style)
    pub table_style_info: Option<TableStyleInfo>,

    // Attributes	Description
    /// comment (Table Comment)
    ///
    /// A string representing a textual comment about the table.
    pub comment: Option<String>,

    /// connectionId (Connection ID)
    ///
    /// An integer representing an ID to indicate which connection from the connections collection is used by this table.
    /// This shall only be used for tables that are based off of xml maps.
    pub connection_id: Option<u64>,

    /// dataCellStyle (Data Style Name)
    ///
    /// A string representing the name of the cell style that is applied to the data area cells of the table.
    /// If this string is missing or does not correspond to the name of a cell style, then the data cell style specified by the current table style should be applied.
    pub data_cell_style: Option<String>,

    /// dataDxfId (Data Area Format Id)
    ///
    /// A zero based integer index into the differential formatting records <dxfs> in the styleSheet indicating which format to apply to the data area of this table.
    /// The spreadsheet should fail to load if this index is out of bounds.
    pub data_dxf_id: Option<u64>,

    /// displayName (Table Name)
    ///
    /// A string representing the name of the table.
    /// This is the name that shall be used in formula references, and displayed in the UI to the spreadsheet user.
    pub display_name: Option<String>,

    /// headerRowBorderDxfId (Header Row Border Format Id)
    ///
    /// A zero based integer index into the differential formatting records <dxfs> in the styleSheet indicating what border formatting to apply to the header row of this table.    /// The spreadsheet should fail to load if this index is out of bounds.
    pub header_row_border_dxf_id: Option<u64>,

    /// headerRowCellStyle (Header Row Style)
    ///
    /// A string representing the name of the cell style that is applied to the header row cells of the table.
    /// If this string is missing or does not correspond to the name of a cell style, then the header row style specified by the current table style should be applied.
    pub header_row_cell_style: Option<String>,

    /// headerRowCount (Header Row Count)
    ///
    /// An integer representing the number of header rows showing at the top of the table. 0 means that the header row is not shown.
    /// It is up to the spreadsheet application to determine if numbers greater than 1 are allowed.
    /// Unless the spreadsheet application has a feature where there might ever be more than one header row, this number should not be higher than 1.
    pub header_row_count: Option<u64>,

    /// headerRowDxfId (Header Row Format Id)
    ///
    /// A zero based integer index into the differential formatting records *<dxfs>*in the styleSheet indicating which format to apply to the header row of this table.
    /// The spreadsheet should fail to load if this index is out of bounds.
    pub header_row_dxf_id: Option<u64>,

    /// id (Table Id)
    ///
    /// A non zero integer representing the unique identifier for this table.
    /// Each table in the workbook shall have a unique id.
    /// Ids can be used to refer to the specific table in the workbook.
    pub id: Option<u64>,

    /// insertRow (Insert Row Showing)
    ///
    /// A Boolean value indicating whether the insert row is showing.
    /// True when the insert row is showing, false otherwise.
    ///
    /// The insert row should only be shown if the table has no data.
    pub insert_row: Option<bool>,

    /// insertRowShift (Insert Row Shift)
    ///
    /// A Boolean that indicates whether cells in the sheet had to be inserted when the insert row was shown for this table.
    /// True if the cells were shifted, false otherwise.
    pub insert_row_shift: Option<bool>,

    /// name (Name)
    ///
    /// A string representing the name of the table that is used to reference the table programmatically through the spreadsheet applications object model.
    /// This string shall be unique per table per sheet.
    ///
    /// By default this should be the same as the table's displayName.
    /// This name should also be kept in synch with the displayName when the displayName is updated in the UI by the spreadsheet user.
    pub name: Option<String>,

    /// published (Published)
    ///
    /// A Boolean representing whether this table is marked as published for viewing by a server based spreadsheet application.
    /// True if it should be viewed by the server spreadsheet application, false otherwise.
    pub published: Option<bool>,

    /// ref (Reference)
    ///
    /// The range on the relevant sheet that the table occupies expressed using A1 style referencing.
    /// The reference shall include the totals row if it is shown.
    pub r#ref: Option<Dimension>,

    /// tableBorderDxfId (Table Border Format Id)
    ///
    /// A zero based integer index into the differential formatting records <dxfs> in the styleSheet indicating what border formatting to apply to the borders of this table.
    pub table_border_dxf_id: Option<u64>,

    /// tableType (Table Type)
    ///
    /// An optional enumeration specifying the type or source of the table.
    /// Indicates whether the table is based off of an external data query, data in a worksheet, or from an xml data mapped to a worksheet.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablevalues?view=openxml-3.0.1
    pub table_type: Option<String>,

    /// totalsRowBorderDxfId (Totals Row Border Format Id)
    ///
    /// A zero based integer index into the differential formatting records <dxfs> in the styleSheet indicating what border formatting to apply to the totals row of this table.
    pub totals_row_border_dxf_id: Option<u64>,

    /// totalsRowCellStyle (Totals Row Style)
    ///
    /// A string representing the name of the cell style that is applied to the totals row cells of the table.
    /// If this string is missing or does not correspond to the name of a cell style, then the totals row style specified by the current table style should be applied.
    pub totals_row_cell_style: Option<String>,

    /// totalsRowCount (Totals Row Count)
    ///
    /// An integer representing the number of totals rows that shall be shown at the bottom of the table.
    /// 0 means that the totals row is not shown.
    /// It is up to the spreadsheet application to determine if numbers greater than 1 are allowed. Unless the spreadsheet application has a feature where their might ever be more than one totals row, this number should not be higher than 1.
    pub totals_row_count: Option<u64>,

    /// totalsRowDxfId (Totals Row Format Id)
    ///
    /// A zero based integer index into the differential formatting records <dxfs> in the styleSheet indicating which format to apply to the totals row of this table.
    pub totals_row_dxf_id: Option<u64>,

    /// totalsRowShown (Totals Row Shown)
    ///
    /// A Boolean indicating whether the totals row has ever been shown in the past for this table.
    /// True if the totals row has been shown, false otherwise.
    pub totals_row_shown: Option<bool>,
}

impl Table {
    pub(crate) fn load(zip: &mut ZipArchive<impl Read + Seek>, path: &str) -> anyhow::Result<Self> {
        let mut table = Self {
            auto_filter: None,
            sort_state: None,
            table_columns: None,
            table_style_info: None,
            comment: None,
            connection_id: None,
            data_cell_style: None,
            data_dxf_id: None,
            display_name: None,
            header_row_border_dxf_id: None,
            header_row_cell_style: None,
            header_row_count: None,
            header_row_dxf_id: None,
            id: None,
            insert_row: None,
            insert_row_shift: None,
            name: None,
            published: None,
            r#ref: None,
            table_border_dxf_id: None,
            table_type: None,
            totals_row_border_dxf_id: None,
            totals_row_cell_style: None,
            totals_row_count: None,
            totals_row_dxf_id: None,
            totals_row_shown: None,
        };

        let Some(mut reader) = xml_reader(zip, path) else {
            return Ok(table);
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"table" => {
                    let attributes = e.attributes();
                    for a in attributes {
                        match a {
                            Ok(a) => {
                                let string_value = String::from_utf8(a.value.to_vec())?;
                                match a.key.local_name().as_ref() {
                                    b"comment" => {
                                        table.comment = Some(string_value);
                                    }
                                    b"connectionId" => {
                                        table.connection_id = string_to_unsignedint(&string_value);
                                    }
                                    b"dataCellStyle" => {
                                        table.data_cell_style = Some(string_value);
                                    }
                                    b"dataDxfId" => {
                                        table.data_dxf_id = string_to_unsignedint(&string_value);
                                    }
                                    b"displayName" => {
                                        table.display_name = Some(string_value);
                                    }
                                    b"headerRowBorderDxfId" => {
                                        table.header_row_border_dxf_id =
                                            string_to_unsignedint(&string_value);
                                    }
                                    b"headerRowCellStyle" => {
                                        table.header_row_cell_style = Some(string_value);
                                    }
                                    b"headerRowCount" => {
                                        table.header_row_count =
                                            string_to_unsignedint(&string_value);
                                    }
                                    b"headerRowDxfId" => {
                                        table.header_row_dxf_id =
                                            string_to_unsignedint(&string_value);
                                    }
                                    b"id" => {
                                        table.id = string_to_unsignedint(&string_value);
                                    }
                                    b"insertRow" => {
                                        table.insert_row = string_to_bool(&string_value);
                                    }
                                    b"insertRowShift" => {
                                        table.insert_row_shift = string_to_bool(&string_value);
                                    }
                                    b"name" => {
                                        table.name = Some(string_value);
                                    }
                                    b"published" => {
                                        table.published = string_to_bool(&string_value);
                                    }
                                    b"ref" => {
                                        let value = a.value.as_ref();
                                        let dimension = a1_dimension_to_row_col(value)?;
                                        table.r#ref = Some(Dimension {
                                            start: Coordinate::from_point(dimension.0),
                                            end: Coordinate::from_point(dimension.1),
                                        });
                                    }
                                    b"tableBorderDxfId" => {
                                        table.table_border_dxf_id =
                                            string_to_unsignedint(&string_value);
                                    }
                                    b"tableType" => {
                                        table.table_type = Some(string_value);
                                    }
                                    b"totalsRowBorderDxfId" => {
                                        table.totals_row_border_dxf_id =
                                            string_to_unsignedint(&string_value);
                                    }
                                    b"totalsRowCellStyle" => {
                                        table.totals_row_cell_style = Some(string_value);
                                    }
                                    b"totalsRowCount" => {
                                        table.totals_row_count =
                                            string_to_unsignedint(&string_value);
                                    }
                                    b"totalsRowDxfId" => {
                                        table.totals_row_dxf_id =
                                            string_to_unsignedint(&string_value);
                                    }
                                    b"totalsRowShown" => {
                                        table.totals_row_shown = string_to_bool(&string_value);
                                    }
                                    _ => {}
                                }
                            }
                            Err(error) => {
                                bail!(error.to_string())
                            }
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"autoFilter" => {
                    table.auto_filter = Some(AutoFilter::load(&mut reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sortState" => {
                    table.sort_state = Some(SortState::load(&mut reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableColumns" => {
                    table.table_columns = Some(load_table_columns(&mut reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableStyleInfo" => {
                    table.table_style_info = Some(TableStyleInfo::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"table" => break,
                Ok(Event::Eof) => break,
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(table);
    }
}
