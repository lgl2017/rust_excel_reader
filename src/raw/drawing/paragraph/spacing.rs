use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use super::{spacing_percent::SpacingPercent, spacing_points::SpacingPoints};

#[derive(Debug, Clone, PartialEq)]
pub enum SpacingEnum {
    // Child Elements
    // spcPct (Spacing Percent)	§21.1.2.2.11
    SpacingPercent(SpacingPercent),
    // spcPts (Spacing Points)
    SpacingPoints(SpacingPoints),
}

impl SpacingEnum {
    pub(crate) fn load(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Option<Self>> {
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    return SpacingEnum::load_helper(e);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(None)
    }

    fn load_helper(e: &BytesStart) -> anyhow::Result<Option<Self>> {
        match e.local_name().as_ref() {
            b"spcPct" => {
                let spacing = SpacingPercent::load(e)?;
                return Ok(Some(SpacingEnum::SpacingPercent(spacing)));
            }
            b"spcPts" => {
                let spacing = SpacingPoints::load(e)?;
                return Ok(Some(SpacingEnum::SpacingPoints(spacing)));
            }
            _ => return Ok(None),
        }
    }
}
