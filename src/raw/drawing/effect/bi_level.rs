use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bilevel?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxBiLevel {
    // attributes
    /// Specifies the luminance threshold for the Bi-Level effect.
    /// Values greater than or equal to the threshold are set to white. Values lesser than the threshold are set to black.
    pub thresh: Option<i64>,
}

impl XlsxBiLevel {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut level = Self { thresh: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"thresh" => {
                            level.thresh = string_to_int(&string_value);
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

        Ok(level)
    }
}
