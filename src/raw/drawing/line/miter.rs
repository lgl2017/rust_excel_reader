use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_unsignedint, raw::drawing::st_types::STPositivePercentage};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.miter?view=openxml-3.0.1
///
/// specifies that a line join shall be mitered
///
/// Example:
/// ```
/// <miter lim="100000" />
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxMiter {
    // Attributes
    /// Specifies the amount by which lines is extended to form a miter join
    ///
    /// lim (Miter Join Limit)
    pub lim: Option<STPositivePercentage>,
}

impl XlsxMiter {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        let mut dash_stop = Self { lim: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"lim" => {
                            dash_stop.lim = string_to_unsignedint(&string_value);
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

        Ok(dash_stop)
    }
}
