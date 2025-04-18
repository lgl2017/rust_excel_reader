use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphareplace?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAlphaReplace {
    // attributes
    /// Specifies the new opacity value.
    pub a: Option<i64>,
}

impl XlsxAlphaReplace {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut effect = Self { a: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"a" => {
                            effect.a = string_to_int(&string_value);
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

        Ok(effect)
    }
}
