#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::common_types::rectangle_alignment::RectangleAlignmentValues,
    raw::drawing::{
        effect::reflection::XlsxReflection,
        st_types::{emu_to_pt, st_angle_to_degree, st_percentage_to_float},
    },
};

/// reflection (Reflection Effect)
///
/// This element specifies a reflection effect.
///
/// Example:
/// ```
/// <a:reflection blurRad="151308" stA="88815" endPos="65000" dist="402621"dir="5400000" sy="-100000" algn="bl" rotWithShape="0" />
/// ```
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.reflection?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Reflection {
    /// aignment
    ///
    /// Specifies shadow alignment; alignment happens first, effectively setting the origin for scale, skew, and offset.
    ///
    /// * Bottom
    /// * BottomLeft
    /// * BottomRight
    /// * Center
    /// * Left
    /// * Right
    /// * Top
    /// * TopLeft
    /// * TopRight
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectanglealignmentvalues?view=openxml-3.0.1
    pub alignment: RectangleAlignmentValues,

    /// blur radius in points
    pub blur_radius: f64,

    /// Specifies the direction to offset the shadow as angle
    pub direction: f64,

    /// Specifies how far to offset the shadow in points
    pub distance: f64,

    /// Specifies the ending reflection opacity in percentage
    pub end_alpha: f64,

    /// Specifies the end position (along the alpha gradient ramp) of the end alpha value as percentage
    pub end_position: f64,

    /// Specifies the direction to offset the reflection in angle degree
    pub fade_direction: f64,

    /// Specifies the horizontal skew angle
    ///
    /// Example: 60 -> 60 degress
    pub horizontal_skew_angle: f64,

    /// Specifies the vertical skew angle
    pub vertical_skew_angle: f64,

    /// Specifies whether the shadow rotates with the shape if the shape is rotated
    pub rotatae_with_shape: bool,

    /// starting reflection opacity in percentage
    // stA (Start Opacity)
    pub start_alpha: f64,

    /// Specifies the start position (along the alpha gradient ramp) of the start alpha value as percentage
    pub start_position: f64,

    /// Specifies the horizontal scaling factor; negative scaling causes a flip in percentage.
    ///
    /// Example: 0.65 -> 65%
    pub horizontal_scale: f64,

    /// Specifies the vertical scaling factor; negative scaling causes a flip.
    pub vertical_scale: f64,
}

impl Reflection {
    pub(crate) fn from_raw(raw: Option<XlsxReflection>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            alignment: RectangleAlignmentValues::from_string(raw.clone().algn),
            blur_radius: emu_to_pt(raw.clone().blur_rad.unwrap_or(0) as i64),
            direction: st_angle_to_degree(raw.clone().dir.unwrap_or(0) as i64),
            distance: emu_to_pt(raw.clone().dist.unwrap_or(0) as i64),
            end_alpha: st_percentage_to_float(raw.clone().end_a.unwrap_or(0) as i64),
            end_position: st_percentage_to_float(raw.clone().end_pos.unwrap_or(0) as i64),
            fade_direction: st_angle_to_degree(raw.clone().fade_dir.unwrap_or(0) as i64),
            horizontal_skew_angle: st_angle_to_degree(raw.clone().kx.unwrap_or(0)),
            vertical_skew_angle: st_angle_to_degree(raw.clone().ky.unwrap_or(0)),
            rotatae_with_shape: raw.clone().rot_with_shape.unwrap_or(false),
            start_alpha: st_percentage_to_float(raw.clone().st_a.unwrap_or(0) as i64),
            start_position: st_percentage_to_float(raw.clone().st_pos.unwrap_or(0) as i64),
            horizontal_scale: st_percentage_to_float(raw.clone().sx.unwrap_or(0)),
            vertical_scale: st_percentage_to_float(raw.clone().sy.unwrap_or(0)),
        });
    }
}
