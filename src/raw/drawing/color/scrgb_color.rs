use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use crate::helper::string_to_int;

use super::color_transforms::ColorTransform;

/// RgbColorModelPercentage: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rgbcolormodelpercentage?view=openxml-3.0.1
///
/// Example: The following represent the same color
/// ```
/// <a:scrgbClr r="50000" g="50000" b="50000"/>
/// <a:srgbClr val="BCBCBC"/> // hex digits RRGGBB.
/// ```
// tag: scrgbClr
#[derive(Debug, Clone, PartialEq)]
pub struct ScrgbColor {
    // attributes: b, g, r
    pub r: Option<i64>,
    pub g: Option<i64>,
    pub b: Option<i64>,
    // children
    pub color_transforms: Option<Vec<ColorTransform>>,
}

impl ScrgbColor {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
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

        color.color_transforms = Some(ColorTransform::load_list(reader, b"scrgbClr")?);

        Ok(color)
    }
}
