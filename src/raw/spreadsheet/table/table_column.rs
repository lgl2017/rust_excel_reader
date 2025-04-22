use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{excel::XmlReader, helper::string_to_unsignedint};

use super::{
    calculated_column_formula::XlsxCalculatedColumnFormula,
    totals_row_formula::XlsxTotalsRowFormula, xml_column_properties::XlsxXmlColumnProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablecolumns?view=openxml-3.0.1
///
/// An element representing the collection of all table columns for this table.
///
/// Example:
/// ```
/// <tableColumns count="8">
///     <tableColumn id="1" xr3:uid="{6ECE91FA-07FC-463E-8474-CA152C66F7BF}" name="GUY" dataDxfId="153" dataCellStyle="List Items" />
///     <tableColumn id="2" name="Segment" />
///     <tableColumn id="3" name="Country" />
/// </tableColumns>
/// ```
///
/// tableColumns (Table Columns)
pub type XlsxTableColumns = Vec<XlsxTableColumn>;

pub(crate) fn load_table_columns(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxTableColumns> {
    let mut buf = Vec::new();
    let mut columns: XlsxTableColumns = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableColumn" => {
                columns.push(XlsxTableColumn::load(reader, e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tableColumns" => break,
            Ok(Event::Eof) => bail!("unexpected end of file at `tableColumns`."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    return Ok(columns);
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablecolumn?view=openxml-3.0.1
///
/// An element representing a single column for this table.
///
/// Example:
/// ```
/// <tableColumns count="8">
///     <tableColumn id="1" xr3:uid="{6ECE91FA-07FC-463E-8474-CA152C66F7BF}" name="GUY" dataDxfId="153" dataCellStyle="List Items" />
///     <tableColumn id="2" name="Segment" />
///     <tableColumn id="3" name="Country" />
/// </tableColumns>
/// ```
///
/// tableColumn (Table Column)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTableColumn {
    /// extLst (Future Feature Data Storage Area)	Not supported

    /// Child Elements	Subclause
    /// calculatedColumnFormula (Calculated Column Formula)	§18.5.1.1
    pub calculated_column_formula: Option<XlsxCalculatedColumnFormula>,

    /// totalsRowFormula (Totals Row Formula)	§18.5.1.6
    pub totals_row_formula: Option<XlsxTotalsRowFormula>,

    /// xmlColumnPr (XML Column Properties)	§18.5.1.7
    pub xml_column_properties: Option<XlsxXmlColumnProperties>,

    // Attributes
    /// dataCellStyle (Data Area Style Name)
    ///
    /// A string representing the name of the cell style that is applied to the cells in the data area of this table column.
    /// If this string is missing or does not correspond to the name of a cell style, then the data cell style specified by the current table style should be applied.
    /// This cell style should get precedence over the dataCellStyle defined by the table.
    pub data_cell_style: Option<String>,

    /// dataDxfId (Data & Insert Row Format Id)
    ///
    /// A zero based integer index into the differential formatting records <dxfs> in the styleSheet indicating which format to apply to the data area of this column.
    /// This formatting shall also apply to cells on the insert row for this column.
    ///
    /// The spreadsheet should fail to load if this index is out of bounds.
    pub data_dxf_id: Option<u64>,

    /// headerRowCellStyle (Header Row Cell Style)
    ///
    /// A string representing the name of the cell style that is applied to the header row cell of this column.
    /// If this string is missing or does not correspond to the name of a cell style, then header row style specified by the current table style should be applied.
    /// This cell style should get precedence over the headerRowCellStyle defined by the table.
    pub header_row_cell_style: Option<String>,

    /// headerRowDxfId (Header Row Cell Format Id)
    ///
    /// A zero based integer index into the differential formatting records <dxfs> in the styleSheet indicating which format to apply to the header cell of this column.
    pub header_row_dxf_id: Option<u64>,

    /// id (Table Field Id)
    ///
    /// An integer representing the unique identifier of this column.
    /// This shall be unique per table.
    pub id: Option<u64>,

    /// name (Column name)
    ///
    /// A string representing the unique caption of the table column.
    /// This is what shall be displayed in the header row in the UI, and is referenced through functions.
    /// This name shall be unique per table.
    pub name: Option<String>,

    /// queryTableFieldId (Query Table Field Id)
    ///
    /// An integer representing the query table field ID corresponding to this table column.
    /// The relationship between this table and the corresponding query table is expressed in _rels part for this table.
    /// Each queryTableField has a unique id attribute, and this id is what is referenced here.
    pub query_table_field_id: Option<u64>,

    /// totalsRowCellStyle (Totals Row Style Name)
    ///
    /// A string representing the name of the cell style that is applied to the Totals Row cell of this column.
    /// If this string is missing or does not correspond to the name of a cell style, then the totals row cell style specified by the current table style should be applied.

    /// This cell style should get precedence over the totalsRowCellStyle defined by the table.
    pub totals_row_cell_style: Option<String>,

    /// totalsRowDxfId (Totals Row Format Id)
    ///
    /// A zero based integer index into the differential formatting records <dxfs> in the styleSheet indicating which format to apply to the totals row cell of this column.
    /// The spreadsheet shall not load if this index is out of bounds.
    pub totals_row_dxf_id: Option<u64>,

    /// totalsRowFunction (Totals Row Function)
    ///
    /// An enumeration indicating which type of aggregation to show in the totals row cell for this column.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.totalsrowfunctionvalues?view=openxml-3.0.1
    pub totals_row_function: Option<String>,

    /// totalsRowLabel (Totals Row Label)
    ///
    /// A String to show in the totals row cell for this column.
    /// This string shall be ignored unless the totalsRowFunction="none" for this column, in which case it is displayed in the totals row.
    pub totals_row_label: Option<String>,

    /// uniqueName (Unique Name)
    ///
    /// An optional string representing the unique name of the table column.
    /// This string is used to bind the column to a field in a data table, so it should only be used when this table's tableType is queryTable or xml.
    ///
    /// This name shall be unique per table when it is used.
    /// For tables created from xml mappings, by default this should be the same as the name of the column, and should be kept in synch with the name of the column if that name is altered by the spreadsheet application.
    pub unique_name: Option<String>,
}

impl XlsxTableColumn {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut column = Self {
            calculated_column_formula: None,
            totals_row_formula: None,
            xml_column_properties: None,
            data_cell_style: None,
            data_dxf_id: None,
            header_row_cell_style: None,
            header_row_dxf_id: None,
            id: None,
            name: None,
            query_table_field_id: None,
            totals_row_cell_style: None,
            totals_row_dxf_id: None,
            totals_row_function: None,
            totals_row_label: None,
            unique_name: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"dataCellStyle" => {
                            column.data_cell_style = Some(string_value);
                        }
                        b"dataDxfId" => {
                            column.data_dxf_id = string_to_unsignedint(&string_value);
                        }
                        b"headerRowCellStyle" => {
                            column.header_row_cell_style = Some(string_value);
                        }
                        b"headerRowDxfId" => {
                            column.header_row_dxf_id = string_to_unsignedint(&string_value);
                        }
                        b"id" => {
                            column.id = string_to_unsignedint(&string_value);
                        }
                        b"name" => {
                            column.name = Some(string_value);
                        }
                        b"queryTableFieldId" => {
                            column.query_table_field_id = string_to_unsignedint(&string_value);
                        }
                        b"totalsRowCellStyle" => {
                            column.totals_row_cell_style = Some(string_value);
                        }
                        b"totalsRowDxfId" => {
                            column.totals_row_dxf_id = string_to_unsignedint(&string_value);
                        }
                        b"totalsRowFunction" => {
                            column.totals_row_function = Some(string_value);
                        }
                        b"totalsRowLabel" => {
                            column.totals_row_label = Some(string_value);
                        }
                        b"uniqueName" => {
                            column.unique_name = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e))
                    if e.local_name().as_ref() == b"calculatedColumnFormula" =>
                {
                    column.calculated_column_formula =
                        Some(XlsxCalculatedColumnFormula::load(reader, e)?)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"totalsRowFormula" => {
                    column.totals_row_formula = Some(XlsxTotalsRowFormula::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xmlColumnPr" => {
                    column.xml_column_properties = Some(XlsxXmlColumnProperties::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tableColumn" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `tableColumn`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(column)
    }
}
