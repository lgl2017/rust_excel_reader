use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{excel::XmlReader, helper::string_to_bool};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculatedcolumnformula?view=openxml-3.0.1
///
/// Columns in a table can have cells that are calculated, usually based on values in other cells in the table.
/// This element stores the formula that is used to perform the calculation for each cell in this column.
///
/// calculatedColumnFormula (Calculated Column Formula)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCalculatedColumnFormula {
    // text content
    pub formula: String,

    // attribute
    /// array (Array)
    ///
    /// A Boolean value that indicates whether this formula is an array style formula.
    pub array: Option<bool>,
}

impl XlsxCalculatedColumnFormula {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut text = String::new();
        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Text(t)) => text.push_str(&t.unescape()?),
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"calculatedColumnFormula" => {
                    break
                }
                Ok(Event::Eof) => bail!("unexpected end of file at `calculatedColumnFormula`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        let attributes = e.attributes();
        let mut formula = Self {
            formula: text,
            array: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"array" => {
                            formula.array = string_to_bool(&string_value);
                            break;
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
