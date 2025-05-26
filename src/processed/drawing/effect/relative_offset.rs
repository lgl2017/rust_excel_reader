#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{
    effect::relative_offset::XlsxRelativeOffset, st_types::st_percentage_to_float,
};

/// relOff (Relative Offset Effect)
///
/// This element specifies a relative offset effect. Sets up a new origin by offsetting relative to the size of the previous effect.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.relativeoffset?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct RelativeOffset {
    /// Specifies the X offset in percentage
    pub horizontal_offset: f64,

    /// Specifies the Y offset in percentage
    pub vertical_offset: f64,
}

impl RelativeOffset {
    pub(crate) fn from_raw(raw: Option<XlsxRelativeOffset>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            horizontal_offset: st_percentage_to_float(raw.clone().tx.unwrap_or(0)),
            vertical_offset: st_percentage_to_float(raw.clone().ty.unwrap_or(0)),
        });
    }
}
