#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::common_types::position::Position,
    raw::drawing::shape::{path::line_to::XlsxLineTo, shape_guide::XlsxShapeGuide},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineto?view=openxml-3.0.1
///
/// This element specifies the drawing of a straight line from the current pen position to the new point specified.
///
/// This line becomes part of the shape geometry, representing a side of the shape.
/// The coordinate system used when specifying this line is the path coordinate system.
///
/// Example
/// ```
/// <a:lnTo>
///     <a:pt x="2650602" y="1261641"/>
/// </a:lnTo>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct LineTo {
    pub end: Position,
}

impl LineTo {
    pub(crate) fn from_raw(raw: XlsxLineTo, guide_list: Option<Vec<XlsxShapeGuide>>) -> Self {
        return Self {
            end: Position::from_point(raw.point, guide_list),
        };
    }
}
