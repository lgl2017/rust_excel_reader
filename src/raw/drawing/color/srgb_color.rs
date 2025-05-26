use quick_xml::events::BytesStart;
use std::io::Read;

use crate::excel::XmlReader;

use crate::common_types::HexColor;
use crate::helper::{extract_val_attribute, hex_to_rgba, rgba_to_hex};

use super::color_transforms::{apply_color_transformations, XlsxColorTransform};

/// RgbColorModelHex: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rgbcolormodelhex?view=openxml-3.0.1
///
/// Example: The following represent the same color
/// ```
/// <a:scrgbClr r="50000" g="50000" b="50000"/>
/// <a:srgbClr val="BCBCBC"/> // hex digits RRGGBB.
/// ```
///
/// tag: srgbClr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSrgbColor {
    // attributes
    /// The actual color value. Expressed as a sequence of hex digits RRGGBB
    pub val: Option<String>,

    // children
    pub color_transforms: Option<Vec<XlsxColorTransform>>,
}

impl XlsxSrgbColor {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let val = extract_val_attribute(e)?;
        let mut color = Self {
            val,
            color_transforms: None,
        };
        color.color_transforms = Some(XlsxColorTransform::load_list(reader, b"srgbClr")?);

        Ok(color)
    }
}

impl XlsxSrgbColor {
    pub(crate) fn to_hex(&self) -> Option<HexColor> {
        let Some(hex) = self.val.clone() else {
            return None;
        };
        let Ok(mut rgba) = hex_to_rgba(&hex, Some(false)) else {
            return None;
        };

        rgba = apply_color_transformations(rgba, self.color_transforms.clone().unwrap_or(vec![]));

        match rgba_to_hex(rgba, Some(false)) {
            Ok(hex) => return Some(hex),
            Err(_) => return None,
        }
    }
}
