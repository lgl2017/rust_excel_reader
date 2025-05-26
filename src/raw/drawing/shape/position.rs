use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::raw::drawing::st_types::STAdjustCoordinate;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.position?view=openxml-3.0.1
///
/// Specifies a position coordinate within the shape bounding box. It should be noted that this coordinate is placed within the shape bounding box using the transform coordinate system which is also called the shape coordinate system, as it encompasses the entire shape. The width and height for this coordinate system are specified within the ext transform element.
///
/// Example:
///
/// ```
/// <a:pos x="2" y="2"/>
/// ```
// tag:  pos (Shape Position Coordinate)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPosition {
    // Attributes
    /// Specifies the x coordinate for this position coordinate.
    ///
    /// value type: `ST_AdjCoordinate` defined as a union of the following
    /// - `ST_Coordinate` simple type: i64
    /// - `ST_GeomGuideName`: String referencing to a geometry guide name
    // x (X-Coordinate)
    pub x: Option<STAdjustCoordinate>,

    /// Specifies the y coordinate for this position coordinate
    ///
    /// value type: `ST_AdjCoordinate` defined as a union of the following
    /// - `ST_Coordinate` simple type: i64
    /// - `ST_GeomGuideName`: String referencing to a geometry guide name
    pub y: Option<STAdjustCoordinate>,
}

impl XlsxPosition {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut position = Self { x: None, y: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"x" => position.x = Some(STAdjustCoordinate::from_string(&string_value)),
                        b"y" => position.y = Some(STAdjustCoordinate::from_string(&string_value)),
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
