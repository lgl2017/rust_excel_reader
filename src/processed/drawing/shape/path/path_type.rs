#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::shape::{path::XlsxPathTypeEnum, shape_guide::XlsxShapeGuide};

use super::{
    arc_to::ArcTo, cubic_bezier_curve_to::CubicBezierCurveTo, line_to::LineTo, move_to::MoveTo,
    quad_bezier_curve_to::QuadraticBezierCurveTo,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.path?view=openxml-3.0.1
///
/// This element specifies a creation path consisting of a series of moves, lines and curves that when combined forms a geometric shape
///
/// Example
/// ```
///   <a:pathLst>
///     <a:path w="2650602" h="1261641">
///       <a:moveTo>
///         <a:pt x="0" y="1261641"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="2650602" y="1261641"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="1226916" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PathTypeValues {
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.arcto?view=openxml-3.0.1
    ///
    /// It draws an arc with the specified parameters from the current pen position to the new point specified.
    /// An arc is a line that is bent based on the shape of a supposed circle.
    /// The length of this arc is determined by specifying both a start angle and an ending angle that act together to effectively specify an end point for the arc.
    ///
    /// This element specifies the existence of an arc within a shape path.
    Arc(ArcTo),

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.closeshapepath?view=openxml-3.0.1
    ///
    /// This element specifies the ending of a series of lines and curves in the creation path of a custom geometric shape.
    /// When this element is encountered, the generating application should consider the corresponding path closed.
    Close,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.cubicbeziercurveto?view=openxml-3.0.1
    ///
    /// This element specifies to draw a cubic bezier curve along the specified points.
    ///
    /// To specify a cubic bezier curve there needs to be 3 points specified.
    /// The first two are control points used in the cubic bezier calculation and the last is the ending point for the curve.
    CubicBezier(CubicBezierCurveTo),

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineto?view=openxml-3.0.1
    ///
    /// This element specifies the drawing of a straight line from the current pen position to the new point specified.
    ///
    /// This line becomes part of the shape geometry, representing a side of the shape.
    /// The coordinate system used when specifying this line is the path coordinate system.
    Line(LineTo),

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.moveto?view=openxml-3.0.1
    ///
    /// This element specifies a set of new coordinates to move the shape cursor to.
    ///
    /// This does not draw a line or curve to this new position from the old position but simply move the cursor to a new starting position.
    /// It is only when a path drawing element such as lnTo is used that a portion of the path is drawn.
    Move(MoveTo),

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.quadraticbeziercurveto?view=openxml-3.0.1
    ///
    /// This element specifies to draw a quadratic bezier curve along the specified points.
    ///
    /// To specify a quadratic bezier curve there needs to be 2 points specified.
    /// The first is a control point used in the quadratic bezier calculation and the last is the ending point for the curve.
    QuadBezier(QuadraticBezierCurveTo),
}

impl PathTypeValues {
    pub(crate) fn from_raw(raw: XlsxPathTypeEnum, guide_list: Option<Vec<XlsxShapeGuide>>) -> Self {
        return match raw {
            XlsxPathTypeEnum::Arc(path) => Self::Arc(ArcTo::from_raw(path, guide_list.clone())),
            XlsxPathTypeEnum::Close(_) => Self::Close,
            XlsxPathTypeEnum::CubicBezier(path) => {
                Self::CubicBezier(CubicBezierCurveTo::from_raw(path, guide_list.clone()))
            }
            XlsxPathTypeEnum::Line(path) => Self::Line(LineTo::from_raw(path, guide_list.clone())),
            XlsxPathTypeEnum::Move(path) => Self::Move(MoveTo::from_raw(path, guide_list.clone())),
            XlsxPathTypeEnum::QuadBezier(path) => {
                Self::QuadBezier(QuadraticBezierCurveTo::from_raw(path, guide_list.clone()))
            }
        };
    }
}
