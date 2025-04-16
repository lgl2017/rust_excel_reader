use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphamodulationfixed?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct AlphaModulationFixed {
    // Attributes
    /// Specifies the percentage amount to scale the alpha.
    // amt (Amount)
    pub amt: Option<i64>,
}

impl AlphaModulationFixed {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut alpha_modulation_fixed = Self { amt: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"amt" => {
                            alpha_modulation_fixed.amt = string_to_int(&string_value);
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

        Ok(alpha_modulation_fixed)
    }
}
