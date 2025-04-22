use std::io::Read;
use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::common_types::HexColor;
use crate::excel::XmlReader;
use crate::helper::{hsla_to_rgba, percentage_int_to_float, rgba_to_hex, string_to_int};

use super::color_transforms::{apply_color_transformations, XlsxColorTransform};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hslcolor?view=openxml-3.0.1
///
/// Example
/// ```
/// // RRGGBB = (00, 00, 80)
/// <a:hslClr hue="14400000" sat="100.000%" lum="50.000%">
/// ```
/// A perceptual gamma of 2.2 is assumed.
// tag: hslClr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxHslColor {
    // attributes:
    /// Specifies the angular value describing the wavelength. Expressed in 1/60000ths of a degree.
    pub hue: Option<i64>,

    /// Specifies the luminance referring to the lightness or darkness of the color.
    /// Expressed as a percentage with 0% referring to maximal dark (black) and 100% referring to maximal white.
    ///
    /// converted to int: 50.000% -> 50000
    pub lum: Option<i64>,

    /// Specifies the saturation referring to the purity of the hue.
    /// Expressed as a percentage with 0% referring to grey, 100% referring to the purest form of the hue.
    ///
    /// converted to int: 50.000% -> 50000
    pub sat: Option<i64>,

    // children
    pub color_transforms: Option<Vec<XlsxColorTransform>>,
}

impl XlsxHslColor {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut color = Self {
            hue: None,
            lum: None,
            sat: None,
            color_transforms: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"hue" => {
                            color.hue = string_to_int(&string_value);
                        }
                        b"lum" => {
                            color.lum = string_to_int(&string_value);
                        }
                        b"sat" => {
                            color.sat = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        color.color_transforms = Some(XlsxColorTransform::load_list(reader, b"hslClr")?);

        Ok(color)
    }
}

impl XlsxHslColor {
    #[allow(dead_code)]
    pub(crate) fn to_hex(&self) -> Option<HexColor> {
        if self.hue.is_none() || self.sat.is_none() || self.lum.is_none() {
            return None;
        }
        let hue = self.hue.unwrap() as f64 / 600.0;
        let lum = percentage_int_to_float(self.lum.unwrap()) * 100.0;
        let sat = percentage_int_to_float(self.sat.unwrap()) * 100.0;
        let Ok(mut rgba) = hsla_to_rgba((hue, lum, sat, 1.0)) else {
            return None;
        };

        rgba = apply_color_transformations(rgba, self.color_transforms.clone().unwrap_or(vec![]));
        match rgba_to_hex(rgba, Some(false)) {
            Ok(hex) => return Some(hex),
            Err(_) => return None,
        }
    }
}
