use anyhow::bail;
use hsl_color::HslColor;
use preset_color::PresetColor;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use scheme_color::SchemeColor;
use scrgb_color::ScrgbColor;
use srgb_color::SrgbColor;
use system_color::SystemColor;

pub mod color_map;
pub mod color_transforms;
pub mod custom_color;
pub mod hsl_color;
pub mod preset_color;
pub mod scheme_color;
pub mod scrgb_color;
pub mod srgb_color;
pub mod system_color;

#[derive(Debug, Clone, PartialEq)]
pub enum ColorEnum {
    // hslClr
    HslColor(HslColor),
    // prstClr
    PresetColor(PresetColor),
    // schemeClr
    SchemeColor(SchemeColor),
    // scrgbClr
    ScrgbColor(ScrgbColor),
    // srgbClr
    SrgbColor(SrgbColor),
    // sysClr
    SystemColor(SystemColor),
}

impl ColorEnum {
    pub(crate) fn load(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Option<Self>> {
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    return ColorEnum::load_helper(reader, e);
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
        let mut colors: Vec<ColorEnum> = vec![];
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if let Some(color) = ColorEnum::load_helper(reader, e)? {
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

    fn load_helper(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Option<Self>> {
        match e.local_name().as_ref() {
            b"hslClr" => {
                let hsl = HslColor::load(reader, e)?;
                return Ok(Some(ColorEnum::HslColor(hsl)));
            }
            b"prstClr" => {
                let preset = PresetColor::load(reader, e)?;
                return Ok(Some(ColorEnum::PresetColor(preset)));
            }
            b"schemeClr" => {
                let scheme = SchemeColor::load(reader, e)?;
                return Ok(Some(ColorEnum::SchemeColor(scheme)));
            }
            b"scrgbClr" => {
                let scrgb: ScrgbColor = ScrgbColor::load(reader, e)?;
                return Ok(Some(ColorEnum::ScrgbColor(scrgb)));
            }
            b"srgbClr" => {
                let srgb = SrgbColor::load(reader, e)?;
                return Ok(Some(ColorEnum::SrgbColor(srgb)));
            }
            b"sysClr" => {
                let system = SystemColor::load(reader, e)?;
                return Ok(Some(ColorEnum::SystemColor(system)));
            }
            _ => return Ok(None),
        }
    }
}

// impl ColorEnum {
//     pub(crate) fn to_hex(&self) {
//         match self {
//             /// <a:hslClr hue="14400000" sat="100.000%" lum="50.000%">
//             ColorEnum::HslColor(hsl_color) => todo!(),
//             ColorEnum::PresetColor(preset_color) => todo!(),
//             ColorEnum::SchemeColor(scheme_color) => todo!(),
//             ColorEnum::ScrgbColor(scrgb_color) => todo!(),
//             ColorEnum::SrgbColor(srgb_color) => todo!(),
//             ColorEnum::SystemColor(system_color) => todo!(),
//         }
//     }
// }
