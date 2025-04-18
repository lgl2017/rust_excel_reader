use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    common_types::{Coordinate, Dimension},
    excel::XmlReader,
    helper::{string_to_bool, string_to_unsignedint},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformula?view=openxml-3.0.1
///
/// Formula for the cell.
/// The formula expression is contained in the character node of this element.
///
/// Example
/// ```
/// <c r="C6" s="1" vm="15">
///     <f>CUBEVALUE("xlextdat9 Adventure Works",C$5,$A6)</f>
///     <v>2838512.355</v>
/// </c>
/// ```
///
/// f (Formula)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCellFormula {
    pub raw_value: String,

    //  attributes
    /// aca (Always Calculate Array)
    ///
    /// Only applies to array formulas.
    /// true indicates that the entire array shall be calculated in full.
    /// If false the individual cells of the array shall be calculated as needed.
    /// The aca value shall be ignored unless the value of the corresponding t attribute is array.
    pub always_calculate_array: Option<bool>,

    /// bx (Assigns Value to Name)
    ///
    /// Specifies that this formula assigns a value to a name.
    pub assign_value_to_name: Option<bool>,

    /// ca (Calculate Cell)
    ///
    /// Indicates that this formula needs to be recalculated the next time calculation is performed.
    pub recalculate_cell: Option<bool>,

    /// del1 (Input 1 Deleted)
    ///
    /// Whether the first input cell for data table has been deleted.
    /// Applies to data table formula only.
    /// Written on master cell of data table formula only.
    pub input_1_deleted: Option<bool>,

    /// del2 (Input 2 Deleted)
    ///
    /// Whether the second input cell for data table has been deleted.
    /// Applies to data table formula only.
    /// Written on master cell of data table formula only.
    pub input_2_deleted: Option<bool>,

    /// dt2D (Data Table 2-D)
    ///
    /// Data table is two-dimentional.
    /// Only applies to the data tables function.
    /// Written on master cell of data table formula only.
    pub data_table_2d: Option<bool>,

    /// dtr (Data Table Row)
    ///
    /// true if one-dimentional data table is a row, otherwise it's a column.
    /// Only applies to the data tables function.
    /// Written on master cell of data table formula only.
    pub data_table_row: Option<bool>,

    /// r1 (Data Table Cell 1)
    ///
    /// First input cell for data table.
    /// Only applies to the data tables array function "TABLE()".
    /// Written on master cell of data table formula only.
    ///
    /// allowed value: ST_CellRef: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/db11a912-b1cb-4dff-b46d-9bedfd10cef0.
    ///
    /// Example: `A1` – a reference to the cell at column 1, row 1.
    pub data_table_cell1: Option<Coordinate>,

    /// r2 (Input Cell 2)
    ///
    /// Second input cell for data table when dt2D is '1'.
    /// Only applies to the data tables array function "TABLE()".
    /// Written on master cell of data table formula only.
    pub data_table_cell2: Option<Coordinate>,

    /// ref (Range of Cells)
    ///
    /// Range of cells which the formula applies to.
    /// Only required for shared formula, array formula or data table.
    /// Only written on the master formula, not subsequent formulas belonging to the same shared group, array, or data table.
    ///
    /// allowed value: ST_Ref: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e7f22870-88a1-4c06-8e5f-d035b1179c50
    ///
    /// Example: `A1:AD5` – a reference to the cell range of 150 cells, with top left corner in the cell at column , row 1, and the bottom right corner in the cell at column 30, row 5. The width of this range is 30 columns, and the height of this range is 5 rows.
    pub ref_range: Option<Dimension>,

    /// si (Shared Group Index)
    ///
    /// Optional attribute to optimize load performance by sharing formulas.
    /// When a formula is a shared formula (t value is shared) then this value indicates the group to which this particular cell's formula belongs.
    /// The first formula in a group of shared formulas is saved in the f element.
    /// This is considered the 'master' formula cell.
    /// Subsequent cells sharing this formula need not have the formula written in their f element.
    /// Instead, the attribute si value for a particular cell is used to figure what the formula expression should be based on the cell's relative location to the master formula cell.
    ///
    /// A cell is shared only when si is used and t is shared.
    /// The formula expression for a cell that is specified to be part of a shared formula (and is not the master) shall be ignored, and the master formula shall override.
    ///
    /// If a master cell of a shared formula range specifies that a particular cell is part of the shared formula range, and that particular cell does not use the si and t attributes to indicate that it is shared, then the particular cell's formula shall override the shared master formula.
    /// If this cell occurs in the middle of a range of shared formula cells, the earlier and later formulas shall continue sharing the master formula, and the cell in question shall not share the formula of the master cell formula.
    ///
    /// Loading and handling of a cell and formula using an si attribute and whose t value is shared, located outside the range specified in the master cell associated with the si group, is implementation defined.
    ///
    /// Master cell references on the same sheet shall not overlap with each other.
    pub shared_group_index: Option<u64>,

    /// t (Formula Type)
    ///
    /// Type of formula.
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformulavalues?view=openxml-3.0.1.
    pub r#type: Option<String>,
}

impl XlsxCellFormula {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut text = String::new();
        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Text(t)) => text.push_str(&t.unescape()?),
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"v" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `v`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        let mut formula = Self {
            raw_value: text,
            always_calculate_array: None,
            assign_value_to_name: None,
            recalculate_cell: None,
            input_1_deleted: None,
            input_2_deleted: None,
            data_table_2d: None,
            data_table_row: None,
            data_table_cell1: None,
            data_table_cell2: None,
            ref_range: None,
            shared_group_index: None,
            r#type: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"aca" => {
                            formula.always_calculate_array = string_to_bool(&string_value);
                        }
                        b"bx" => {
                            formula.assign_value_to_name = string_to_bool(&string_value);
                        }
                        b"ca" => {
                            formula.recalculate_cell = string_to_bool(&string_value);
                        }
                        b"del1" => {
                            formula.input_1_deleted = string_to_bool(&string_value);
                        }
                        b"del2" => {
                            formula.input_2_deleted = string_to_bool(&string_value);
                        }
                        b"dt2D" => {
                            formula.data_table_2d = string_to_bool(&string_value);
                        }
                        b"dtr" => {
                            formula.data_table_row = string_to_bool(&string_value);
                        }
                        b"r1" => {
                            let value = a.value.as_ref();
                            formula.data_table_cell1 = Coordinate::from_a1(value);
                        }
                        b"r2" => {
                            let value = a.value.as_ref();
                            formula.data_table_cell2 = Coordinate::from_a1(value);
                        }
                        b"ref" => {
                            let value = a.value.as_ref();
                            formula.ref_range = Dimension::from_a1(value);
                        }
                        b"si" => {
                            formula.shared_group_index = string_to_unsignedint(&string_value);
                        }
                        b"t" => {
                            formula.r#type = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(formula)
    }
}
