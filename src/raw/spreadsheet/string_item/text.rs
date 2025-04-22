use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

use crate::{common_types::Text, excel::XmlReader};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.text?view=openxml-3.0.1
///
/// This element represents the text content shown as part of a string.
///
/// Example:
/// ```
/// <si>
///     <t>some string</t>
/// </si>
/// ```
// tag: t
pub(crate) fn load_text(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Text> {
    let mut text = String::new();
    let mut buf: Vec<u8> = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Text(t)) => text.push_str(&t.unescape()?),
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"t" => break,
            Ok(Event::Eof) => bail!("unexpected end of file at `t`."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    return Ok(text);
}
