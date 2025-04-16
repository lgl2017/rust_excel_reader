use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tableparts?view=openxml-3.0.1
///
/// This collection expresses a relationship Id pointing to every table on this sheet.
///
/// Example
/// ```
/// <tableParts count="2">
///   <tablePart r:id="rId1"/>
///   <tablePart r:id="rId2"/>
/// </tableParts>
/// ```
pub type TableParts = Vec<TablePart>;

pub(crate) fn load_table_parts(reader: &mut XmlReader) -> anyhow::Result<TableParts> {
    let mut parts: TableParts = vec![];

    let mut buf = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tablePart" => {
                parts.push(TablePart::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tableParts" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(parts)
}
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablepart?view=openxml-3.0.1
///
/// A single Table Part reference.
///
/// Example:
/// ```
/// <tablePart r:id="rId2" />
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct TablePart {
    // attributes
    /// This relationship Id is used to locate a particular table definition part.
    pub id: String,
}

impl TablePart {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"id" => return Ok(Self { id: string_value }),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        bail!("id not specified for the table part.")
    }
}
