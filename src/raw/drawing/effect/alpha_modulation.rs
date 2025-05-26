use super::effect_container::XlsxEffectContainer;
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

/// Alpha Modulate Effect: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphamodulationeffect?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAlphaModulation {
    // children

    // cont (Effect Container)
    pub cont: Option<Box<XlsxEffectContainer>>,
}

impl XlsxAlphaModulation {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut modulation = Self { cont: None };
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cont" => {
                    modulation.cont = Some(Box::new(XlsxEffectContainer::load(reader, e)?));
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"alphaMod" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        return Ok(modulation);
    }
}
