use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_int, raw::drawing::st_types::STCoordinate};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.childextents?view=openxml-3.0.1
///
/// This element specifies the size dimensions of the child extents rectangle and is used for calculations of grouping, scaling, and rotation behavior of shapes placed within a group.
///
/// Example:
/// ```
/// <a:chExt cx="2426208" cy="978408"/>
/// ```
pub type XlsxChildExtent = XlsxExtents;

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
pub struct XlsxExtents {
    // attributes
    /// Extent Length
    // tag: cx
    pub cx: Option<STCoordinate>,

    /// Extent Width
    // tag: cy
    pub cy: Option<STCoordinate>,
}

impl XlsxExtents {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut position = Self { cx: None, cy: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"cx" => position.cx = string_to_int(&string_value),
                        b"cy" => position.cy = string_to_int(&string_value),
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
