#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{line::custom_dash::XlsxDashStop, st_types::st_percentage_to_float};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.dashstop?view=openxml-3.0.1
///
/// This element specifies a dash stop primitive.
/// Dashing schemes are built by specifying an ordered list of dash stop primitive.
/// A dash stop primitive consists of a dash and a space.
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct DashStop {
    /// Specifies the length of the dash relative to the line width in percentage
    ///
    /// Example: 65% -> 0.65
    pub dash_length: f64,

    /// Specifies the length of the space relative to the line width in percentage.
    pub space_length: f64,
}

impl DashStop {
    pub(crate) fn from_raw(raw: XlsxDashStop) -> Self {
        return Self {
            dash_length: st_percentage_to_float(raw.d.unwrap_or(0) as i64),
            space_length: st_percentage_to_float(raw.sp.unwrap_or(0) as i64),
        };
    }
}
