#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{effect::blur::XlsxBlur, st_types::emu_to_pt};

/// blur (Blur Effect)
///
/// a blur effect that is applied to the entire shape, including its fill.
/// All color channels, including alpha, are affected.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blur?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Blur {
    /// Specifies whether the bounds of the object should be grown as a result of the blurring.
    ///
    /// True indicates the bounds are grown while false indicates that they are not.
    pub grow: bool,

    /// Specifies the radius of blur: positive number in points
    pub rad: f64,
}

impl Blur {
    pub(crate) fn from_raw(raw: Option<XlsxBlur>) -> Option<Self> {
        let Some(raw) = raw else {
            return None;
        };
        let Some(rad) = raw.rad.clone() else {
            return None;
        };

        return Some(Self {
            grow: raw.grow.unwrap_or(false),
            rad: emu_to_pt(rad as i64),
        });
    }
}
