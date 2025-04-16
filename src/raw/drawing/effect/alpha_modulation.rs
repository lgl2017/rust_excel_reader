use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use super::effect_container::EffectContainer;

/// Alpha Modulate Effect: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphamodulationeffect?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct AlphaModulation {
    // children

    // cont (Effect Container)
    pub cont: Option<Box<EffectContainer>>,
}

impl AlphaModulation {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut modulation = Self { cont: None };
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cont" => {
                    modulation.cont = Some(Box::new(EffectContainer::load(reader, e)?));
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
