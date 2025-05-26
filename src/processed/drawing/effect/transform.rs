#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::effect::transform::XlsxTransformEffect;
use crate::raw::drawing::st_types::{emu_to_pt, st_angle_to_degree, st_percentage_to_float};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transformeffect?view=openxml-3.0.1
///
/// This element specifies a transform effect. The transform is applied to each point in the shape's geometry using the following matrix:
///
/// xfrm (Transform Effect)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct TransformEffect {
    /// Specifies the horizontal skew angle
    ///
    /// Example: 60 -> 60 degress
    pub horizontal_skew_angle: f64,

    /// Specifies the vertical skew angle
    pub vertical_skew_angle: f64,

    /// Specifies the horizontal scaling factor; negative scaling causes a flip in percentage.
    ///
    /// Example: 0.65 -> 65%
    pub horizontal_scale: f64,

    /// Specifies the vertical scaling factor; negative scaling causes a flip.
    pub vertical_scale: f64,

    /// Specifies an amount by which to shift the object along the x-axis in points.
    pub horizontal_shift: f64,

    /// Specifies an amount by which to shift the object along the y-axis in points.
    pub vertical_shift: f64,
}

impl TransformEffect {
    pub(crate) fn from_raw(raw: Option<XlsxTransformEffect>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            horizontal_skew_angle: st_angle_to_degree(raw.clone().kx.unwrap_or(0)),
            vertical_skew_angle: st_angle_to_degree(raw.clone().ky.unwrap_or(0)),
            horizontal_scale: st_percentage_to_float(raw.clone().sx.unwrap_or(0)),
            vertical_scale: st_percentage_to_float(raw.clone().sy.unwrap_or(0)),
            horizontal_shift: emu_to_pt(raw.clone().tx.unwrap_or(0)),
            vertical_shift: emu_to_pt(raw.clone().ty.unwrap_or(0)),
        });
    }
}
