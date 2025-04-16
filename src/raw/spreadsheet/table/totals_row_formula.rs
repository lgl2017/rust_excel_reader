use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{excel::XmlReader, helper::string_to_bool};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.totalsrowformula?view=openxml-3.0.1
///
/// This element contains a custom formula for aggregating values from the column.
/// Each tableColumn has a totalsRowFunction that can be used for simple aggregations such as average, standard deviation, min, max, count, and others.
/// If a more custom calculation is desired, then this element should be used, and the totalsRowFunction shall be set to "custom".
///
/// totalsRowFormula (Totals Row Formula)
#[derive(Debug, Clone, PartialEq)]
pub struct TotalsRowFormula {
    // text content
    pub formula: String,

    // attribute
    /// array (Array)
    ///
    /// A Boolean value that indicates whether this formula is an array style formula.
    pub array: Option<bool>,
}

impl TotalsRowFormula {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut text = String::new();
        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Text(t)) => text.push_str(&t.unescape()?),
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"totalsRowFormula" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `totalsRowFormula`."),
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
