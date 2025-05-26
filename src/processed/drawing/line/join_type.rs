#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{line::outline::XlsxOutline, st_types::st_percentage_to_float};

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LineJoinTypeValue {
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linejoinbevel?view=openxml-3.0.1
    ///
    /// specifies that an angle joint is used to connect lines
    Bevel,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.round?view=openxml-3.0.1
    ///
    /// specifies that a round join is used to connect lines
    Round,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.miter?view=openxml-3.0.1
    ///
    /// specifies that a line join shall be mitered
    ///
    /// Example:
    /// ```
    /// <miter lim="100000" />
    /// // lines are extended for 100% to form a miter join
    /// ```
    ///
    /// associated value:
    /// Specifies the amount by which lines are extended to form a miter join in percentage.
    ///
    /// Example: `1.0` represents `100%`
    Miter(f64),
}

impl LineJoinTypeValue {
    pub(crate) fn default() -> Self {
        Self::Miter(8.0)
    }

    pub(crate) fn from_raw(raw: XlsxOutline, reference: Option<Self>) -> Self {
        if let Some(_) = raw.bevel {
            return Self::Bevel;
        }
        if let Some(_) = raw.round {
            return Self::Round;
        }
        if let Some(join) = raw.miter {
            let lim = st_percentage_to_float(join.lim.unwrap_or(0) as i64);
            return Self::Miter(lim);
        }

        return reference.unwrap_or(Self::default());
    }
}
