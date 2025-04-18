use anyhow::bail;
use cell_formula::XlsxCellFormula;
use cell_value::XlsxCellValue;
use inline_string::{load_inline_string, XlsxInlineString};
use quick_xml::events::{BytesStart, Event};

use crate::{
    common_types::Coordinate,
    excel::XmlReader,
    helper::{string_to_bool, string_to_unsignedint},
};

pub mod cell_formula;
pub mod cell_value;
pub mod inline_string;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1
///
/// This element represents a cell in the worksheet.
///
/// Example:
/// ```
/// // A2 contains the shared string at index 1
/// <c r="A2" t="s" s="6">
///     <v>1</v>
/// </c>
///
/// // C2: empty cell
/// <c r="C2" s="7" />
///
/// // address in the grid is C6, whose style index is '1', and whose value metadata index is '15'.
/// // The cell contains a formula as well as a calculated result of that formula.
/// <c r="C6" s="1" vm="15">
///     <f>CUBEVALUE("xlextdat9 Adventure Works",C$5,$A6)</f>
///     <v>2838512.355</v>
/// </c>
///
/// // While a cell can have a formula element f and a value element v, when the cell's type t is inlineStr then only the element is is allowed as a child element.
/// // expressing a string in the cell rather than using the shared string table
/// <c r="A1" t="inlineStr">
///     <is><t>This is inline string example</t></is>
/// </c>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCell {
    /// extLst (Future Feature Data Storage Area)	Not supporte

    /// Child Elements
    /// f (Formula)	§18.3.1.40
    pub formula: Option<XlsxCellFormula>,

    /// is (Rich Text Inline)	§18.3.1.53
    pub inline_string: Option<XlsxInlineString>,

    /// v (Cell Value)
    pub cell_value: Option<XlsxCellValue>,

    /// Attributes	Description
    /// cm (Cell Metadata Index)
    ///
    /// The zero-based index of the cell metadata record associated with this cell.
    /// Metadata information is found in the Metadata Part.
    /// Cell metadata is extra information stored at the cell level, and is attached to the cell (travels through moves, copy / paste, clear, etc).
    /// Cell metadata is not accessible via formula reference.
    pub cell_metadata: Option<u64>,

    /// ph (Show Phonetic)
    ///
    /// A Boolean value indicating if the spreadsheet application should show phonetic information.
    /// Phonetic information is displayed in the same cell across the top of the cell and serves as a 'hint' which indicates how the text should be pronounced.
    /// This should only be used for East Asian languages.
    pub show_phonetic: Option<bool>,

    /// r (Reference)
    ///
    /// An A1 style reference to the location of this cell, ie: "A1".
    ///
    /// Converted to R1C1
    pub coordinate: Option<Coordinate>,

    /// s (Style Index)
    ///
    /// The index of this cell's style.
    ///
    /// 0 based index reference to `cellXfs` in stylesheet.
    pub style: Option<u64>,

    /// t (Cell Data Type)
    ///
    /// An enumeration representing the cell's data type.
    /// Possible values: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_CellType_topic_ID0E6NEFB.html
    pub r#type: Option<String>,

    /// vm (Value Metadata Index)
    ///
    /// The zero-based index of the value metadata record associated with this cell's value.
    /// Metadata records are stored in the Metadata Part. Value metadata is extra information stored at the cell level, but associated with the value rather than the cell itself. Value metadata is accessible via formula reference.
    pub value_metadata: Option<u64>,
}

impl XlsxCell {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut cell = Self {
            formula: None,
            inline_string: None,
            cell_value: None,
            cell_metadata: None,
            show_phonetic: None,
            coordinate: None,
            style: None,
            r#type: None,
            value_metadata: None,
        };

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"cm" => {
                            cell.cell_metadata = string_to_unsignedint(&string_value);
                        }
                        b"ph" => {
                            cell.show_phonetic = string_to_bool(&string_value);
                        }
                        b"r" => {
                            cell.coordinate = Coordinate::from_a1(&a.value);
                        }
                        b"s" => {
                            cell.style = string_to_unsignedint(&string_value);
                        }
                        b"t" => {
                            cell.r#type = Some(string_value);
                        }
                        b"vm" => {
                            cell.value_metadata = string_to_unsignedint(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"f" => {
                    cell.formula = Some(XlsxCellFormula::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"is" => {
                    cell.inline_string = Some(load_inline_string(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"v" => {
                    cell.cell_value = Some(XlsxCellValue::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `c`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        if cell.coordinate.is_none() {
            bail!("Cell of unknwon position.")
        }

        return Ok(cell);
    }
}
