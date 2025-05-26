#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::common_types::adjust_coordinate::AdjustCoordinate,
    raw::drawing::{
        shape::shape_guide::XlsxShapeGuide, text::shape_text_rectangle::XlsxShapeTextRectangle,
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectangle?view=openxml-3.0.1
///
/// This element specifies the rectangular bounding box for text within a `custGeom` shape.
/// The default for this rectangle is the bounding box for the shape.
/// This can be modified using this elements four attributes to inset or extend the text bounding box.
///
/// Example:
/// ```
/// <a:rect l="0" t="0" r="0" b="0"/>
/// ```
/// rect (Shape Text Rectangle)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]

pub struct ShapeTextRectangle {
    /// t (Top)
    ///
    /// Specifies the y coordinate in points of the top edge for a shape text rectangle.
    pub top: AdjustCoordinate,

    /// b (Bottom Position)
    ///
    /// Specifies the y coordinate in points of the bottom edge for a shape text rectangle.
    pub bottom: AdjustCoordinate,

    /// l (Left)
    ///
    /// Specifies the x coordinate in points of the left edge for a shape text rectangle.
    pub left: AdjustCoordinate,

    /// r (Right)
    ///
    /// Specifies the x coordinate in points of the right edge for a shape text rectangle.
    pub right: AdjustCoordinate,
}

impl ShapeTextRectangle {
    pub(crate) fn default() -> Self {
        Self {
            top: AdjustCoordinate::default(),
            bottom: AdjustCoordinate::default(),
            left: AdjustCoordinate::default(),
            right: AdjustCoordinate::default(),
        }
    }
    pub(crate) fn from_raw(
        raw: Option<XlsxShapeTextRectangle>,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };

        return Self {
            top: AdjustCoordinate::from_raw(raw.t, guide_list.clone()),
            bottom: AdjustCoordinate::from_raw(raw.b, guide_list.clone()),
            left: AdjustCoordinate::from_raw(raw.l, guide_list.clone()),
            right: AdjustCoordinate::from_raw(raw.r, guide_list.clone()),
        };
    }
}
