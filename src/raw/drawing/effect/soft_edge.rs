use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_unsignedint, raw::drawing::st_types::STPositiveCoordinate};

/// softEdge (Soft Edge Effect)
///
/// This element specifies a soft edge effect.
/// The edges of the shape are blurred, while the fill is not affected.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.softedge?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSoftEdge {
    // attributes
    /// Specifies the radius of blur to apply to the edges.
    ///
    /// rad (Radius)
    pub rad: Option<STPositiveCoordinate>,
}

impl XlsxSoftEdge {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut soft_edge = Self { rad: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"rad" => {
                            soft_edge.rad = string_to_unsignedint(&string_value);
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

        Ok(soft_edge)
    }
}
