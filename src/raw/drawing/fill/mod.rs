pub mod blip_fill;
pub mod fill_rectangle;
pub mod fill_reference;
pub mod gradient_fill;
pub mod group_fill;
pub mod no_fill;
pub mod pattern_fill;
pub mod solid_fill;

use anyhow::bail;
use std::io::Read;

use crate::excel::XmlReader;
use blip_fill::XlsxBlipFill;
use gradient_fill::XlsxGradientFill;
use group_fill::XlsxGroupFill;
use no_fill::XlsxNoFill;
use pattern_fill::XlsxPatternFill;
use quick_xml::events::{BytesStart, Event};
use solid_fill::XlsxSolidFill;

/// BackgroundFillStyleList:
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.backgroundfillstylelist?view=openxml-3.0.1
///
/// Defines a set of three background fill styles that are used within a theme, arranged in order from subtle to moderate to intense
///
/// Example
/// ```
/// <a:bgFillStyleLst>
///     <a:solidFill>
///         <a:schemeClr val="phClr" />
///     </a:solidFill>
///     <gradFill rotWithShape="1">
///     …  </gradFill>
///     <blipFill>
///     …  </blipFill>
/// </a:bgFillStyleLst>
/// ```
///
/// tag: bgFillStyleLst
pub type XlsxBackgroundFillStyleList = Vec<XlsxFillStyleEnum>;

pub(crate) fn load_bg_fill_style_lst(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxBackgroundFillStyleList> {
    return Ok(XlsxFillStyleEnum::load_list(reader, b"bgFillStyleLst")?);
}

/// FillStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillstylelist?view=openxml-3.0.1
/// This element defines a set of three fill styles that are used within a theme.
/// The three fill styles are arranged in order from subtle to moderate to intense.
// tag: fillStyleLst
pub type XlsxFillStyleList = Vec<XlsxFillStyleEnum>;

pub(crate) fn load_fill_style_lst(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxFillStyleList> {
    return Ok(XlsxFillStyleEnum::load_list(reader, b"fillStyleLst")?);
}

#[derive(Debug, Clone, PartialEq)]
pub enum XlsxFillStyleEnum {
    // tag: solidFill
    SolidFill(XlsxSolidFill),

    // tag: gradFill
    GradientFill(XlsxGradientFill),

    // grpFill (Group Fill)	§20.1.8.35
    GroupFill(XlsxGroupFill),

    // noFill (No Fill)	§20.1.8.44
    NoFill(XlsxNoFill),

    // pattFill (Pattern Fill)	§20.1.8.47
    PatternFill(XlsxPatternFill),

    // blipFill (Picture Fill)	§20.1.8.14
    BlipFill(XlsxBlipFill),
}

impl XlsxFillStyleEnum {
    pub(crate) fn load(
        reader: &mut XmlReader<impl Read>,
        tag: &[u8],
    ) -> anyhow::Result<Option<Self>> {
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    return XlsxFillStyleEnum::load_helper(reader, e);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(None)
    }

    pub(crate) fn load_list(
        reader: &mut XmlReader<impl Read>,
        tag: &[u8],
    ) -> anyhow::Result<Vec<Self>> {
        let mut fills: Vec<XlsxFillStyleEnum> = vec![];
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if let Some(fill) = XlsxFillStyleEnum::load_helper(reader, e)? {
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

    fn load_helper(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
    ) -> anyhow::Result<Option<Self>> {
        match e.local_name().as_ref() {
            b"solidFill" => {
                let Some(fill) = XlsxSolidFill::load(reader, b"solidFill")? else {
                    return Ok(None);
                };
                return Ok(Some(XlsxFillStyleEnum::SolidFill(fill)));
            }
            b"gradFill" => {
                let fill = XlsxGradientFill::load(reader, e)?;
                return Ok(Some(XlsxFillStyleEnum::GradientFill(fill)));
            }
            b"grpFill" => {
                return Ok(Some(XlsxFillStyleEnum::GroupFill(true)));
            }
            b"noFill" => {
                return Ok(Some(XlsxFillStyleEnum::NoFill(true)));
            }
            b"pattFill" => {
                let fill = XlsxPatternFill::load(reader, e)?;
                return Ok(Some(XlsxFillStyleEnum::PatternFill(fill)));
            }

            b"blipFill" => {
                let fill = XlsxBlipFill::load(reader, e)?;
                return Ok(Some(XlsxFillStyleEnum::BlipFill(fill)));
            }
            _ => return Ok(None),
        }
    }
}
