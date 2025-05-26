#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textverticaloverflowvalues?view=openxml-3.0.1
///
/// Determines whether the text can flow out of the bounding box vertically.
/// If this attribute is omitted, then a value of `overflow` is implied.
///
/// * Clip
/// * Ellipse
/// * Overflow
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextVerticalOverflowValues {
    Clip,
    Ellipsis,
    Overflow,
}

impl TextVerticalOverflowValues {
    pub(crate) fn default() -> Self {
        Self::Overflow
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "clip" => Self::Clip,
            "ellipsis" => Self::Ellipsis,
            "overflow" => Self::Overflow,
            _ => Self::default(),
        };
    }
}
