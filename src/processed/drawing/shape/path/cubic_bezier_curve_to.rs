#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::common_types::position::Position,
    raw::drawing::shape::{
        path::cubic_bezier_curve_to::XlsxCubicBezierCurveTo, shape_guide::XlsxShapeGuide,
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.cubicbeziercurveto?view=openxml-3.0.1
///
/// This element specifies to draw a cubic bezier curve along the specified points.
///
/// To specify a cubic bezier curve there needs to be 3 points specified.
/// The first two are control points used in the cubic bezier calculation and the last is the ending point for the curve.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct CubicBezierCurveTo {
    /// The first control point of the curve.
    pub control1: Position,

    /// The second control point of the curve.
    pub control2: Position,

    /// ending point for the curve
    pub end: Position,
}

impl CubicBezierCurveTo {
    pub(crate) fn from_raw(
        raw: XlsxCubicBezierCurveTo,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        let raw_points = raw.points.unwrap_or(vec![]);
        let count = raw_points.clone().iter().count();

        let control1 = if count > 0 {
            Some(raw_points[0].clone())
        } else {
            None
        };

        let control2 = if count > 1 {
            Some(raw_points[1].clone())
        } else {
            None
        };

        let end = if count > 2 {
            Some(raw_points[2].clone())
        } else {
            None
        };

        return Self {
            control1: Position::from_point(control1, guide_list.clone()),
            control2: Position::from_point(control2, guide_list.clone()),
            end: Position::from_point(end, guide_list.clone()),
        };
    }
}
