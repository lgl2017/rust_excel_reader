#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::common_types::position::Position,
    raw::drawing::shape::{path::move_to::XlsxMoveTo, shape_guide::XlsxShapeGuide},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.moveto?view=openxml-3.0.1
///
/// This element specifies a set of new coordinates to move the shape cursor to.
///
/// This does not draw a line or curve to this new position from the old position but simply move the cursor to a new starting position.
/// It is only when a path drawing element such as lnTo is used that a portion of the path is drawn.
///
/// Example
/// ```
/// <a:moveTo>
///     <a:pt x="0" y="1261641"/>
/// </a:moveTo>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct MoveTo {
    pub end: Position,
}

impl MoveTo {
    pub(crate) fn from_raw(raw: XlsxMoveTo, guide_list: Option<Vec<XlsxShapeGuide>>) -> Self {
        return Self {
            end: Position::from_point(raw.point, guide_list),
        };
    }
}
