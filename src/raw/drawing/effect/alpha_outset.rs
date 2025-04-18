use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphaoutset?view=openxml-3.0.1
/// This is equivalent to an alpha ceiling, followed by alpha blur, followed by either an alpha ceiling (positive radius) or alpha floor (negative radius).
// tag: alphaOutset
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAlphaOutset {
    // attributes
    /// Specifies the radius of blur.
    // tag: rad (Radius)
    pub rad: Option<i64>,
}

impl XlsxAlphaOutset {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut blur = Self { rad: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"rad" => {
                            blur.rad = string_to_int(&string_value);
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

        Ok(blur)
    }
}
