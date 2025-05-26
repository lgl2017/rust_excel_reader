#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{
    effect::hue_saturation_luminance::XlsxHsl,
    st_types::{st_angle_to_degree, st_percentage_to_float},
};

/// hsl (Hue Saturation Luminance Effect)
///
/// This element specifies a hue/saturation/luminance effect.
/// The hue, saturation, and luminance can each be adjusted relative to its current value.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hsl?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct HslEffect {
    /// Specifies the number of degrees by which the hue is adjusted.
    ///
    /// Example: 60 -> 60 degress
    pub hue: f64,

    /// Specifies the percentage by which the luminance is adjusted.
    ///
    /// Example: 0.65 -> 65%
    pub lum: f64,

    /// Specifies the percentage by which the saturation is adjusted.
    ///
    /// Example: 0.65 -> 65%
    pub sat: f64,
}

impl HslEffect {
    pub(crate) fn from_raw(raw: Option<XlsxHsl>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            hue: st_angle_to_degree(raw.clone().hue.unwrap_or(0)),
            lum: st_percentage_to_float(raw.clone().lum.unwrap_or(0)),
            sat: st_percentage_to_float(raw.clone().lum.unwrap_or(0)),
        });
    }
}
