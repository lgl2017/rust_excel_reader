pub mod color_map;
pub mod color_transforms;
pub mod custom_color;
pub mod hsl_color;
pub mod preset_color;
pub mod scheme_color;
pub mod scrgb_color;
pub mod srgb_color;
pub mod system_color;

use anyhow::bail;
use std::io::Read;

use super::scheme::color_scheme::XlsxColorScheme;
use crate::{common_types::HexColor, excel::XmlReader};

use hsl_color::XlsxHslColor;
use preset_color::XlsxPresetColor;
use quick_xml::events::{BytesStart, Event};
use scheme_color::XlsxSchemeColor;
use scrgb_color::XlsxScrgbColor;
use srgb_color::XlsxSrgbColor;
use system_color::XlsxSystemColor;

#[derive(Debug, Clone, PartialEq)]
pub enum XlsxColorEnum {
    // hslClr
    HslColor(XlsxHslColor),
    // prstClr
    PresetColor(XlsxPresetColor),
    // schemeClr
    SchemeColor(XlsxSchemeColor),
    // scrgbClr
    ScrgbColor(XlsxScrgbColor),
    // srgbClr
    SrgbColor(XlsxSrgbColor),
    // sysClr
    SystemColor(XlsxSystemColor),
}

impl XlsxColorEnum {
    pub(crate) fn load(
        reader: &mut XmlReader<impl Read>,
        tag: &[u8],
    ) -> anyhow::Result<Option<Self>> {
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    return XlsxColorEnum::load_helper(reader, e);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(None)
    }

    #[allow(dead_code)]
    pub(crate) fn load_list(
        reader: &mut XmlReader<impl Read>,
        tag: &[u8],
    ) -> anyhow::Result<Vec<Self>> {
        let mut colors: Vec<XlsxColorEnum> = vec![];
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if let Some(color) = XlsxColorEnum::load_helper(reader, e)? {
                        colors.push(color);
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(colors)
    }

    fn load_helper(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
    ) -> anyhow::Result<Option<Self>> {
        match e.local_name().as_ref() {
            b"hslClr" => {
                let hsl = XlsxHslColor::load(reader, e)?;
                return Ok(Some(XlsxColorEnum::HslColor(hsl)));
            }
            b"prstClr" => {
                let preset = XlsxPresetColor::load(reader, e)?;
                return Ok(Some(XlsxColorEnum::PresetColor(preset)));
            }
            b"schemeClr" => {
                let scheme = XlsxSchemeColor::load(reader, e)?;
                return Ok(Some(XlsxColorEnum::SchemeColor(scheme)));
            }
            b"scrgbClr" => {
                let scrgb = XlsxScrgbColor::load(reader, e)?;
                return Ok(Some(XlsxColorEnum::ScrgbColor(scrgb)));
            }
            b"srgbClr" => {
                let srgb = XlsxSrgbColor::load(reader, e)?;
                return Ok(Some(XlsxColorEnum::SrgbColor(srgb)));
            }
            b"sysClr" => {
                let system = XlsxSystemColor::load(reader, e)?;
                return Ok(Some(XlsxColorEnum::SystemColor(system)));
            }
            _ => return Ok(None),
        }
    }
}

impl XlsxColorEnum {
    #[allow(dead_code)]
    pub(crate) fn to_hex(
        &self,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<HexColor> {
        match self {
            XlsxColorEnum::HslColor(c) => c.to_hex(),
            XlsxColorEnum::PresetColor(c) => c.to_hex(),
            XlsxColorEnum::SchemeColor(c) => c.to_hex(color_scheme.clone(), ref_color),
            XlsxColorEnum::ScrgbColor(c) => c.to_hex(),
            XlsxColorEnum::SrgbColor(c) => c.to_hex(),
            XlsxColorEnum::SystemColor(c) => c.to_hex(),
        }
    }
}
