use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{common_types::Text, excel::XmlReader, helper::string_to_unsignedint};

use super::text::load_text;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.phoneticrun?view=openxml-3.0.1
///
/// This element represents a run of text which displays a phonetic hint for this String Item (si).
///
/// Example
/// ```
/// <si>
///     <t>課きく　毛こ</t>
///     <rPh sb="0" eb="1">
///         <t>カ</t>
///     </rPh>
///     <rPh sb="4" eb="5">
///        <t>ケ</t>
///     </rPh>
///     <phoneticPr fontId="1"/>
/// </si>
/// ```
///
/// The above example shows a String Item that displays some Japanese text "課きく　毛こ."
/// It also displays some phonetic text across the top of the cell.
/// The phonetic text character, "カ" is displayed over the "課" character and the phonetic text "ケ" is displayed above the "毛" character, using the font record in the style sheet at index 1.
// tag: rPh
#[derive(Debug, Clone, PartialEq)]
pub struct PhoneticRun {
    // child
    // t (Text)
    pub text: Option<Text>,

    // attributes
    /// An integer used as a zero-based index representing the ending offset into the base text for this phonetic run.
    /// This represents the ending point in the base text the phonetic hint applies to.
    /// This value shall be between 0 and the total length of the base text.
    /// The following condition shall be true: sb < eb.
    // eb (Base Text End Index)
    pub base_text_end_index: Option<u64>,

    /// An integer used as a zero-based index representing the starting offset into the base text for this phonetic run.
    /// This represents the starting point in the base text the phonetic hint applies to.
    // sb (Base Text Start Index)
    pub base_text_start_index: Option<u64>,
}

impl PhoneticRun {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut run = Self {
            text: None,
            base_text_end_index: None,
            base_text_start_index: None,
        };

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"eb" => {
                            run.base_text_end_index = string_to_unsignedint(&string_value);
                        }
                        b"sb" => {
                            run.base_text_start_index = string_to_unsignedint(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" => {
                    run.text = Some(load_text(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"rPh" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `rPh`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        return Ok(run);
    }
}
