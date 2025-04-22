use std::io::Read;
use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use crate::common_types::HexColor;
use crate::helper::{extract_val_attribute, hex_to_rgba, rgba_to_hex};

use super::color_transforms::{apply_color_transformations, XlsxColorTransform};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.systemcolor?view=openxml-3.0.1
///
/// Example: The following represent the same color
/// ```
/// <a:sysClr val="window" lastClr="FFFFFF" />
/// ```
// tag: sysClr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSystemColor {
    // attributes:

    // tag: lastClr
    pub last_clr: Option<String>,

    /// allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.systemcolorvalues?view=openxml-3.0.1
    pub val: Option<String>,

    // children
    pub color_transforms: Option<Vec<XlsxColorTransform>>,
}

impl XlsxSystemColor {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let val = extract_val_attribute(e)?;
        let mut color = Self {
            last_clr: None,
            val,
            color_transforms: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"lastClr" => {
                            color.last_clr = Some(string_value);
                            break;
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        color.color_transforms = Some(XlsxColorTransform::load_list(reader, b"sysClr")?);
        Ok(color)
    }
}

impl XlsxSystemColor {
    pub(crate) fn to_hex(&self) -> Option<HexColor> {
        if self.last_clr.is_none() {
            return None;
        }
        let hex = self.last_clr.clone().unwrap();
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
