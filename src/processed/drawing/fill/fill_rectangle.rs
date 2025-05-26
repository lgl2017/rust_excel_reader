#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::fill::fill_rectangle::XlsxFillRectangle;
use crate::raw::drawing::st_types::st_percentage_to_float;

/// specifies a fill rectangle of an image in relation to the bounding box
///
/// Each edge of the fill rectangle is defined by a percentage offset from the corresponding edge of the shape's bounding box.
/// A positive percentage specifies an inset, while a negative percentage specifies an outset.
/// For example, a left offset of 25% specifies that the left edge of the fill rectangle is located to the right of the bounding box's left edge by an amount equal to 25% of the bounding box's width.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillrectangle?view=openxml-3.0.1
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, PartialOrd, Clone)]
pub struct FillRectangle {
    /// Specifies offset from the bottom edge of the rectangle in percentage.
    pub bottom_offset: f64,

    /// Specifies offset from the left edge of the rectangle in percentage.
    pub left_offset: f64,

    /// Specifies offset from the right edge of the rectangle in percentage.
    pub right_offset: f64,

    /// Specifies offset from the top edge of the rectangle in percentage.
    pub top_offset: f64,
}

impl FillRectangle {
    pub(crate) fn default() -> Self {
        return Self {
            bottom_offset: 0.0,
            left_offset: 0.0,
            right_offset: 0.0,
            top_offset: 0.0,
        };
    }

    pub(crate) fn from_raw(raw_rect: Option<XlsxFillRectangle>) -> Self {
        let Some(raw_rect) = raw_rect else {
            return Self::default();
        };

        return Self {
            bottom_offset: st_percentage_to_float(raw_rect.b.unwrap_or(0)),
            left_offset: st_percentage_to_float(raw_rect.l.unwrap_or(0)),
            right_offset: st_percentage_to_float(raw_rect.r.unwrap_or(0)),
            top_offset: st_percentage_to_float(raw_rect.t.unwrap_or(0)),
        };
    }
}
