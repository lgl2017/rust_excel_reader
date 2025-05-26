#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{effect::luminance::XlsxLuminance, st_types::st_percentage_to_float};

/// lum (Luminance Effect)
///
/// This element specifies a luminance effect.
/// Brightness linearly shifts all colors closer to white or black.
/// Contrast scales all colors to be either closer or further apart.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminanceeffect?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Luminance {
    /// Specifies the percent to change the brightness.
    ///
    /// Example: 0.65 -> 65%
    pub bright: f64,

    /// Specifies the percent to change the contrast
    ///
    /// Example: 0.65 -> 65%
    pub contrast: f64,
}

impl Luminance {
    pub(crate) fn from_raw(raw: Option<XlsxLuminance>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            bright: st_percentage_to_float(raw.clone().bright.unwrap_or(0)),
            contrast: st_percentage_to_float(raw.clone().contrast.unwrap_or(0)),
        });
    }
}
