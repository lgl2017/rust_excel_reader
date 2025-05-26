use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use super::effect_container::XlsxEffectContainer;

/// blend (Blend Effect)
///
/// Specifies a blend of several effects.
/// The container specifies the raw effects to blend while the blend mode specifies how the effects are to be blended.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blend?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxBlend {
    // children
    /// specifies the raw effects to blend
    pub cont: Option<Box<XlsxEffectContainer>>,

    // attributes
    /// Specifies how to blend the two effects
    /// Allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blendmodevalues?view=openxml-3.0.1
    pub blend: Option<String>,
}

impl XlsxBlend {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut blend = Self {
            cont: None,
            blend: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"blend" => {
                            blend.blend = Some(string_value);
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
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cont" => {
                    blend.cont = Some(Box::new(XlsxEffectContainer::load(reader, e)?));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"blend" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(blend)
    }
}
