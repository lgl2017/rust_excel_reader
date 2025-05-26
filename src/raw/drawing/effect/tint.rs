use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

/// tint (Tint Effect)
///
/// This element specifies a tint effect.
/// Shifts effect color values towards/away from hue by the specified amount.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tinteffect?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTint {
    // attributes:
    /// Specifies by how much the color value is shifted.
    pub amt: Option<i64>,

    /// Specifies the hue towards which to tint.
    pub hue: Option<i64>,
}

impl XlsxTint {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut tint = Self {
            amt: None,
            hue: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"amt" => {
                            tint.amt = string_to_int(&string_value);
                        }
                        b"hue" => {
                            tint.hue = string_to_int(&string_value);
                        }

                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(tint)
    }
}
