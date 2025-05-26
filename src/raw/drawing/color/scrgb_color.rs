use anyhow::bail;
use quick_xml::events::BytesStart;
use std::io::Read;

use crate::common_types::HexColor;
use crate::excel::XmlReader;

use crate::helper::{rgba_to_hex, string_to_int};
use crate::raw::drawing::st_types::{st_percentage_to_float, STPercentage};

use super::color_transforms::{apply_color_transformations, XlsxColorTransform};

/// scrgbClr (RgbColorModelPercentage)
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rgbcolormodelpercentage?view=openxml-3.0.1
///
/// Example: The following represent the same color
/// ```
/// <a:scrgbClr r="50000" g="50000" b="50000"/>
/// <a:srgbClr val="BCBCBC"/> // hex digits RRGGBB.
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxScrgbColor {
    // attributes: b, g, r
    pub r: Option<STPercentage>,
    pub g: Option<STPercentage>,
    pub b: Option<STPercentage>,

    // children
    pub color_transforms: Option<Vec<XlsxColorTransform>>,
}

impl XlsxScrgbColor {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut color = Self {
            r: None,
            g: None,
            b: None,
            color_transforms: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"r" => {
                            color.r = string_to_int(&string_value);
                        }
                        b"g" => {
                            color.g = string_to_int(&string_value);
                        }
                        b"b" => {
                            color.b = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        color.color_transforms = Some(XlsxColorTransform::load_list(reader, b"scrgbClr")?);

        Ok(color)
    }
}

impl XlsxScrgbColor {
    pub(crate) fn to_hex(&self) -> Option<HexColor> {
        if self.r.is_none() || self.g.is_none() || self.b.is_none() {
            return None;
        }

        let r = st_percentage_to_float(self.r.unwrap()) * 255.0;
        let g = st_percentage_to_float(self.g.unwrap()) * 255.0;
        let b = st_percentage_to_float(self.b.unwrap()) * 255.0;

        let mut rgba: (u32, u32, u32, f64) =
            (r.round() as u32, g.round() as u32, b.round() as u32, 1.0);

        rgba = apply_color_transformations(rgba, self.color_transforms.clone().unwrap_or(vec![]));

        match rgba_to_hex(rgba, Some(false)) {
            Ok(hex) => return Some(hex),
            Err(_) => return None,
        }
    }
}
