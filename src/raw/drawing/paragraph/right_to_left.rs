use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_bool;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.righttoleft?view=openxml-3.0.1
///
/// This element specifies whether the contents of this run shall have right-to-left characteristics.
/// True when this element’s present without `val` attribute or the `val` attribute is true
// tag: rtl
#[derive(Debug, Clone, PartialEq)]
pub struct RightToLeft {
    pub val: Option<bool>,
}

impl RightToLeft {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut rtl = Self { val: Some(true) };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"val" => {
                            rtl.val = string_to_bool(&string_value);
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
        Ok(rtl)
    }
}
