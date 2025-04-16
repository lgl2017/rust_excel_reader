use anyhow::bail;
use blip_fill::BlipFill;
use gradient_fill::GradientFill;
use group_fill::GroupFill;
use no_fill::NoFill;
use pattern_fill::PatternFill;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use solid_fill::SolidFill;

pub mod blip_fill;
pub mod fill_reference;
pub mod gradient_fill;
pub mod group_fill;
pub mod no_fill;
pub mod pattern_fill;
pub mod rectangle;
pub mod solid_fill;

/// BackgroundFillStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.backgroundfillstylelist?view=openxml-3.0.1
/// defines a set of three background fill styles that are used within a theme, arranged in order from subtle to moderate to intense
///
/// Example
/// ````
/// <a:bgFillStyleLst>
///     <a:solidFill>
///         <a:schemeClr val="phClr" />
///     </a:solidFill>
///     <gradFill rotWithShape="1">
///     …  </gradFill>
///     <blipFill>
///     …  </blipFill>
/// </a:bgFillStyleLst>
/// ``
// tag: bgFillStyleLst
pub type BackgroundFillStyleList = Vec<FillStyleEnum>;

pub(crate) fn load_bg_fill_style_lst(
    reader: &mut XmlReader,
) -> anyhow::Result<BackgroundFillStyleList> {
    return Ok(FillStyleEnum::load_list(reader, b"bgFillStyleLst")?);
}

/// FillStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillstylelist?view=openxml-3.0.1
/// This element defines a set of three fill styles that are used within a theme.
/// The three fill styles are arranged in order from subtle to moderate to intense.
// tag: fillStyleLst
pub type FillStyleList = Vec<FillStyleEnum>;

pub(crate) fn load_fill_style_lst(reader: &mut XmlReader) -> anyhow::Result<FillStyleList> {
    return Ok(FillStyleEnum::load_list(reader, b"fillStyleLst")?);
}

#[derive(Debug, Clone, PartialEq)]
pub enum FillStyleEnum {
    // tag: solidFill
    SolidFill(SolidFill),

    // tag: gradFill
    GradientFill(GradientFill),

    // grpFill (Group Fill)	§20.1.8.35
    GroupFill(GroupFill),

    // noFill (No Fill)	§20.1.8.44
    NoFill(NoFill),

    // pattFill (Pattern Fill)	§20.1.8.47
    PatternFill(PatternFill),

    // blipFill (Picture Fill)	§20.1.8.14
    BlipFill(BlipFill),
}

impl FillStyleEnum {
    pub(crate) fn load(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Option<Self>> {
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    return FillStyleEnum::load_helper(reader, e);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(None)
    }

    pub(crate) fn load_list(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Vec<Self>> {
        let mut fills: Vec<FillStyleEnum> = vec![];
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if let Some(fill) = FillStyleEnum::load_helper(reader, e)? {
                        fills.push(fill);
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(fills)
    }

    fn load_helper(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Option<Self>> {
        match e.local_name().as_ref() {
            b"solidFill" => {
                let Some(fill) = SolidFill::load(reader, b"solidFill")? else {
                    return Ok(None);
                };
                return Ok(Some(FillStyleEnum::SolidFill(fill)));
            }
            b"gradFill" => {
                let fill = GradientFill::load(reader, e)?;
                return Ok(Some(FillStyleEnum::GradientFill(fill)));
            }
            b"grpFill" => {
                return Ok(Some(FillStyleEnum::GroupFill(true)));
            }
            b"noFill" => {
                return Ok(Some(FillStyleEnum::NoFill(true)));
            }
            b"pattFill" => {
                let fill = PatternFill::load(reader, e)?;
                return Ok(Some(FillStyleEnum::PatternFill(fill)));
            }

            b"blipFill" => {
                let fill = BlipFill::load(reader, e)?;
                return Ok(Some(FillStyleEnum::BlipFill(fill)));
            }
            _ => return Ok(None),
        }
    }
}
