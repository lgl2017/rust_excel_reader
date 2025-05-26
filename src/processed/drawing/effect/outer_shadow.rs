#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::scheme::color_scheme::XlsxColorScheme;
use crate::raw::drawing::st_types::{emu_to_pt, st_angle_to_degree, st_percentage_to_float};
use crate::{
    common_types::HexColor,
    processed::drawing::common_types::rectangle_alignment::RectangleAlignmentValues,
    raw::drawing::effect::outer_shadow::XlsxOuterShadow,
};

/// outerShdw (Outer Shadow Effect)
///
/// specifies an outer shadow effect.
///
///  Example:
/// ```
/// <a:outerShdw blurRad="57150" dist="38100" dir="5400000" algn="ctr" rotWithShape="0" >
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:outerShdw>
/// ```
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outershadow?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct OuterShadow {
    /// color
    pub color: HexColor,

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

    /// Specifies the horizontal skew angle
    ///
    /// Example: 60 -> 60 degress
    pub horizontal_skew_angle: f64,

    /// Specifies the vertical skew angle
    pub vertical_skew_angle: f64,

    /// Specifies whether the shadow rotates with the shape if the shape is rotated
    pub rotatae_with_shape: bool,

    /// Specifies the horizontal scaling factor; negative scaling causes a flip in percentage.
    ///
    /// Example: 0.65 -> 65%
    pub horizontal_scale: f64,

    /// Specifies the vertical scaling factor; negative scaling causes a flip.
    pub vertical_scale: f64,
}

impl OuterShadow {
    pub(crate) fn from_raw(
        raw: Option<XlsxOuterShadow>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };

        let Some(color) = raw.clone().color else {
            return None;
        };
        let Some(hex) = color.to_hex(color_scheme.clone(), ref_color.clone()) else {
            return None;
        };

        return Some(Self {
            color: hex,
            alignment: RectangleAlignmentValues::from_string(raw.clone().algn),
            blur_radius: emu_to_pt(raw.clone().blur_rad.unwrap_or(0) as i64),
            direction: st_angle_to_degree(raw.clone().dir.unwrap_or(0) as i64),
            distance: emu_to_pt(raw.clone().dist.unwrap_or(0) as i64),
            horizontal_skew_angle: st_angle_to_degree(raw.clone().kx.unwrap_or(0)),
            vertical_skew_angle: st_angle_to_degree(raw.clone().ky.unwrap_or(0)),
            rotatae_with_shape: raw.clone().rot_with_shape.unwrap_or(false),
            horizontal_scale: st_percentage_to_float(raw.clone().sx.unwrap_or(0)),
            vertical_scale: st_percentage_to_float(raw.clone().sy.unwrap_or(0)),
        });
    }
}
