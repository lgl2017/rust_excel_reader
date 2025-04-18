use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::common_types::XlsxAdjustCoordinate;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.offset?view=openxml-3.0.1
///
/// This element specifies the location of the bounding box of an object
///
/// Example
/// ```
/// <a:xfrm>
///   <a:off x="3200400" y="1600200"/>
///   <a:ext cx="1705233" cy="679622"/>
/// </a:xfrm>
/// ```
// tag: off
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxOffset {
    // Attributes
    /// X-Axis Coordinate.
    // x (X-Coordinate)
    pub x: Option<XlsxAdjustCoordinate>,

    /// Y-Axis Coordinate
    // y (y-Coordinate)
    pub y: Option<XlsxAdjustCoordinate>,
}

impl XlsxOffset {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut offset = Self { x: None, y: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"x" => offset.x = Some(XlsxAdjustCoordinate::from_string(&string_value)),
                        b"y" => offset.y = Some(XlsxAdjustCoordinate::from_string(&string_value)),
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
