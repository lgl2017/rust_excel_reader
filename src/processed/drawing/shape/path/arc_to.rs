#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::common_types::{
        adjust_angle::AdjustAngle, adjust_coordinate::AdjustCoordinate,
    },
    raw::drawing::shape::{path::arc_to::XlsxArcTo, shape_guide::XlsxShapeGuide},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.arcto?view=openxml-3.0.1
///
/// This element specifies the existence of an arc within a shape path.
///
/// It draws an arc with the specified parameters from the current pen position to the new point specified.
/// An arc is a line that is bent based on the shape of a supposed circle.
/// The length of this arc is determined by specifying both a start angle and an ending angle that act together to effectively specify an end point for the arc.
///
/// Example
/// ```
///   <a:pathLst>
///     <a:path w="2650602" h="1261641">
///       <a:arcTo hR="123"/>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ArcTo {
    /// hR (Shape Arc Height Radius)
    ///
    /// This attribute specifies the height radius of the supposed circle being used to draw the arc.
    /// This gives the circle a total height of (2 * hR).
    /// This total height could also be called it's vertical diameter as it is the diameter for the y axis only.
    pub height_radius: AdjustCoordinate,

    /// wR (Shape Arc Width Radius)
    ///
    /// This attribute specifies the width radius of the supposed circle being used to draw the arc.
    /// This gives the circle a total width of (2 * wR).
    /// This total width could also be called it's horizontal diameter as it is the diameter for the x axis only.
    pub width_radius: AdjustCoordinate,

    /// stAng (Shape Arc Start Angle)
    ///
    /// Specifies the start angle for an arc.
    /// This angle specifies what angle along the supposed circle path is used as the start position for drawing the arc.
    /// This start angle is locked to the last known pen position in the shape path. Thus guaranteeing a continuos shape path.
    pub start_angle: AdjustAngle,

    /// swAng (Shape Arc Swing Angle)
    ///
    /// Specifies the swing angle for an arc.
    /// This angle specifies how far angle-wise along the supposed cicle path the arc is extended.
    /// The extension from the start angle is always in the clockwise direction around the supposed circle.
    pub swing_angle: AdjustAngle,
}

impl ArcTo {
    pub(crate) fn from_raw(raw: XlsxArcTo, guide_list: Option<Vec<XlsxShapeGuide>>) -> Self {
        return Self {
            height_radius: AdjustCoordinate::from_raw(raw.height_radius, guide_list.clone()),
            width_radius: AdjustCoordinate::from_raw(raw.width_radius, guide_list.clone()),
            start_angle: AdjustAngle::from_raw(raw.start_angle, guide_list.clone()),
            swing_angle: AdjustAngle::from_raw(raw.swing_angle, guide_list.clone()),
        };
    }
}
