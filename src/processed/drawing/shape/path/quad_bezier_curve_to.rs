#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::common_types::position::Position,
    raw::drawing::shape::{
        path::quad_bezier_curve_to::XlsxQuadraticBezierCurveTo, shape_guide::XlsxShapeGuide,
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.quadraticbeziercurveto?view=openxml-3.0.1
///
/// This element specifies to draw a quadratic bezier curve along the specified points.
///
/// To specify a quadratic bezier curve there needs to be 2 points specified.
/// The first is a control point used in the quadratic bezier calculation and the last is the ending point for the curve.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct QuadraticBezierCurveTo {
    /// control point used in the quadratic bezier calculation
    pub control: Position,

    /// ending point for the curve
    pub end: Position,
}

impl QuadraticBezierCurveTo {
    pub(crate) fn from_raw(
        raw: XlsxQuadraticBezierCurveTo,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        let raw_points = raw.points.unwrap_or(vec![]);
        let count = raw_points.clone().iter().count();

        let control = if count > 0 {
            Some(raw_points[0].clone())
        } else {
            None
        };

        let end = if count > 2 {
            Some(raw_points[2].clone())
        } else {
            None
        };

        return Self {
            control: Position::from_point(control, guide_list.clone()),
            end: Position::from_point(end, guide_list.clone()),
        };
    }
}
