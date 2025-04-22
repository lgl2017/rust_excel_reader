use std::io::Read;
use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use super::rotation::XlsxRotation;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lightrig?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:lightRig rig="twoPt" dir="t">
///     <a:rot lat="0" lon="0" rev="6000000"/>
/// </a:lightRig>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxLightRig {
    // children
    pub rot: Option<XlsxRotation>,

    // attributes
    /// Direction
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lightrigdirectionvalues?view=openxml-3.0.1
    pub dir: Option<String>,

    /// Rig Preset
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lightrigvalues?view=openxml-3.0.1
    pub rig: Option<String>,
}

impl XlsxLightRig {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut rig = Self {
            rot: None,
            dir: None,
            rig: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"dir" => {
                            rig.dir = Some(string_value);
                        }
                        b"rig" => {
                            rig.rig = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rot" => {
                    rig.rot = Some(XlsxRotation::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"lightRig" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(rig)
    }
}
