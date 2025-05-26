#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{
    st_types::{st_percentage_to_float, st_text_point_to_pt},
    text::paragraph::spacing::XlsxSpacingEnum,
};

/// Type used for Line Spacing, Spacing Before, Spacing After.
///
/// * [Percentage](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpercent?view=openxml-3.0.1)
///     - This element specifies the amount of white space that is to be used between lines and paragraphs in the form of a percentage of the text size.
///
/// * [Point](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpoints?view=openxml-3.0.1)
///     - This element specifies the amount of white space that is to be used between lines and paragraphs in the form of a text point size.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum SpacingTypeValues {
    Percentage(f64),
    Point(f64),
}

impl SpacingTypeValues {
    pub(crate) fn defualt() -> Self {
        Self::Point(0.0)
    }
    pub(crate) fn from_raw(raw: Option<XlsxSpacingEnum>) -> Self {
        let Some(raw) = raw else {
            return Self::defualt();
        };

        return match raw {
            XlsxSpacingEnum::SpacingPercent(percent) => {
                Self::Percentage(st_percentage_to_float(percent.val.unwrap_or(0) as i64))
            }
            XlsxSpacingEnum::SpacingPoints(pt) => {
                Self::Point(st_text_point_to_pt(pt.val.unwrap_or(0) as i64))
            }
        };
    }
}
