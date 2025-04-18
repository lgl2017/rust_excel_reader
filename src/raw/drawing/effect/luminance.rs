use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminanceeffect?view=openxml-3.0.1
// tag: lum
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxLuminance {
    // attributes:
    /// Specifies the percent to change the brightness.
    pub bright: Option<i64>,

    /// Specifies the percent to change the contrast
    pub contrast: Option<i64>,
}

impl XlsxLuminance {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut lum = Self {
            bright: None,
            contrast: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"bright" => {
                            lum.bright = string_to_int(&string_value);
                        }
                        b"contrast" => {
                            lum.contrast = string_to_int(&string_value);
                        }

                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(lum)
    }
}
