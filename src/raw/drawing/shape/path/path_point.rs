use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::common_types::XlsxAdjustCoordinate;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.point?view=openxml-3.0.1
///
/// This element specifies an x-y coordinate within the path coordinate space
///
/// Example:
///
/// ```
/// <a:pt x="1226916" y="0"/>
/// ```
// tag:  pos (Shape Position Coordinate)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPoint {
    // Attributes
    /// Specifies the x coordinate for this position coordinate.
    // x (X-Coordinate)
    pub x: Option<XlsxAdjustCoordinate>,

    /// Specifies the y coordinate for this position coordinate
    // y (y-Coordinate)
    pub y: Option<XlsxAdjustCoordinate>,
}

impl XlsxPoint {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut point = Self { x: None, y: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"x" => point.x = Some(XlsxAdjustCoordinate::from_string(&string_value)),
                        b"y" => point.y = Some(XlsxAdjustCoordinate::from_string(&string_value)),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(point)
    }
}
