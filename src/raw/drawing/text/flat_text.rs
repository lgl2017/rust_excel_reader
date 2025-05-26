use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_int, raw::drawing::st_types::STCoordinate};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.flattext?view=openxml-3.0.1
///
/// Keep text out of 3D scene entirely
// tag: flatTx
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFlatText {
    // Attributes
    /// z (Z Coordinate)
    ///
    /// Specifies the Z coordinate to be used in positioning the flat text within the 3D scene.
    pub z: Option<STCoordinate>,
}

impl XlsxFlatText {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut flat_text = Self { z: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"z" => {
                            flat_text.z = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(flat_text)
    }
}
