use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_int, raw::drawing::st_types::STPercentage};

/// relOff (Relative Offset Effect)
///
/// This element specifies a relative offset effect. Sets up a new origin by offsetting relative to the size of the previous effect.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.relativeoffset?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxRelativeOffset {
    // attributes
    /// Specifies the X offset in percentage
    pub tx: Option<STPercentage>,

    /// Specifies the Y offset in percentage
    pub ty: Option<STPercentage>,
}

impl XlsxRelativeOffset {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut offset = Self { tx: None, ty: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"tx" => {
                            offset.tx = string_to_int(&string_value);
                        }
                        b"ty" => {
                            offset.ty = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(offset)
    }
}
