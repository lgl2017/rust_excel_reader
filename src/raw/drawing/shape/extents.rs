use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.extents?view=openxml-3.0.1
///
/// This element specifies the size of the bounding box enclosing the referenced object.
///
/// Example
/// ```
/// <a:xfrm>
///   <a:off x="3200400" y="1600200"/>
///   <a:ext cx="1705233" cy="679622"/>
/// </a:xfrm>
/// ```
// tag: ext
#[derive(Debug, Clone, PartialEq)]
pub struct Extents {
    // attributes
    /// Extent Length
    // tag: cx
    pub length: Option<i64>,

    /// Extent Width
    // tag: cy
    pub width: Option<i64>,
}

impl Extents {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut position = Self {
            length: None,
            width: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"cx" => position.length = string_to_int(&string_value),
                        b"cy" => position.width = string_to_int(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(position)
    }
}
