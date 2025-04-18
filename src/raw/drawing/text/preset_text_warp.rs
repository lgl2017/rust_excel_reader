use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    raw::drawing::shape::adjust_value_list::{load_adjust_value_list, XlsxAdjustValueList},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presettextwarp?view=openxml-3.0.1
///
/// Example
/// ```
/// <prstTxWarp prst="textArchUp">
///     <a:avLst>
///         <a:gd name="myGuide" fmla="val 2"/>
///     </a:avLst>
/// </prstTxWarp>
/// ```
// tag: prstTxWarp
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPresetTextWarp {
    // children
    /// Adjust Value List
    // tag: avLst
    pub adjust_value_list: Option<XlsxAdjustValueList>,

    // attributes
    /// Preset Warp Shape
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textshapevalues?view=openxml-3.0.1
    pub preset: Option<String>,
}

impl XlsxPresetTextWarp {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut preset_text_warp = Self {
            adjust_value_list: None,
            preset: None,
        };

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"prst" => {
                            preset_text_warp.preset = Some(string_value);
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

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"avLst" => {
                    preset_text_warp.adjust_value_list = Some(load_adjust_value_list(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"prstTxWarp" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(preset_text_warp)
    }
}
