#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectanglealignmentvalues?view=openxml-3.0.1
///
/// * Bottom
/// * BottomLeft
/// * BottomRight
/// * Center
/// * Left
/// * Right
/// * Top
/// * TopLeft
/// * TopRight
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum RectangleAlignmentValues {
    Bottom,
    BottomLeft,
    BottomRight,
    Center,
    Left,
    Right,
    Top,
    TopLeft,
    TopRight,
}

impl RectangleAlignmentValues {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::Center };
        return match s.as_ref() {
            "b" => Self::Bottom,
            "bl" => Self::BottomLeft,
            "br" => Self::BottomRight,
            "ctr" => Self::Center,
            "l" => Self::Left,
            "r" => Self::Right,
            "t" => Self::Top,
            "tl" => Self::TopLeft,
            "tr" => Self::TopRight,
            _ => Self::Center,
        };
    }
}
