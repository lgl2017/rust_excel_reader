#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{
    effect::tint::XlsxTint,
    st_types::{st_angle_to_degree, st_percentage_to_float},
};

/// tint (Tint Effect)
///
/// This element specifies a tint effect.
/// Shifts effect color values towards/away from hue by the specified amount.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tinteffect?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Tint {
    /// Specifies by how much the color value is shifted as percentage.
    ///
    /// Example: 0.65 -> 65%
    pub amount: f64,

    /// Specifies the hue towards which to tint in angle.
    ///
    /// Example: 60 -> 60 degress
    pub hue: f64,
}

impl Tint {
    pub(crate) fn from_raw(raw: Option<XlsxTint>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            hue: st_angle_to_degree(raw.clone().hue.unwrap_or(0)),
            amount: st_percentage_to_float(raw.clone().amt.unwrap_or(0)),
        });
    }
}
