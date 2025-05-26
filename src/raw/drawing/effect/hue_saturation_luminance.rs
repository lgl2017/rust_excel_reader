use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

/// hsl (Hue Saturation Luminance Effect)
///
/// This element specifies a hue/saturation/luminance effect.
/// The hue, saturation, and luminance can each be adjusted relative to its current value.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hsl?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxHsl {
    // attributes:
    /// Specifies the number of degrees by which the hue is adjusted.
    pub hue: Option<i64>,

    /// Specifies the percentage by which the luminance is adjusted.
    pub lum: Option<i64>,

    /// Specifies the percentage by which the saturation is adjusted.
    pub sat: Option<i64>,
}

impl XlsxHsl {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut hsl = Self {
            hue: None,
            lum: None,
            sat: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"hue" => {
                            hsl.hue = string_to_int(&string_value);
                        }
                        b"lum" => {
                            hsl.lum = string_to_int(&string_value);
                        }
                        b"sat" => {
                            hsl.sat = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(hsl)
    }
}
