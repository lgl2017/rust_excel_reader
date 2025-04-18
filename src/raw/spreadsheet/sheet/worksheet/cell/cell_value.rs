use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalue?view=openxml-3.0.1
///
/// This element expresses the value contained in a cell.
/// If the cell contains a string, then this value is an 0 based index into the shared string table, pointing to the actual string value.
/// Otherwise, the value of the cell is expressed directly in this element.
/// Cells containing formulas express the last calculated result of the formula in this element.
///
/// Example
/// ```
/// <c r="C6" s="1" vm="15">
///     <f>CUBEVALUE("xlextdat9 Adventure Works",C$5,$A6)</f>
///     <v>2838512.355</v>
/// </c>
/// // B4 contains the number "360"
/// <c r="B4">
///     <v>360</v>
/// </c>
/// // C4 contains the UTC date 22 November 1976, 08:30
/// <c r="C4" t="d">
///     <v>1976-11-22T08:30Z</v>
/// </c>
/// // A2 contains the shared string at index 1
/// <c r="A2" t="s" s="6">
///     <v>1</v>
/// </c>
/// ```
///
/// v (Cell value)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCellValue {
    pub raw_value: String,

    // attributes
    /// Content Contains Significant Whitespace.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spaceprocessingmodevalues?view=openxml-3.0.1
    pub space: Option<String>,
}

impl XlsxCellValue {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut space: Option<String> = None;

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"space" => {
                            space = Some(string_value);
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

        Ok(Self {
            raw_value: text,
            space,
        })
    }
}
