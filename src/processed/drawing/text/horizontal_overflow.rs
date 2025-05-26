#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.texthorizontaloverflowvalues?view=openxml-3.0.1
///
/// Determines whether the text can flow out of the bounding box horizontally.
///
/// This is used to determine what will happen in the event that the text within a shape is too large for the bounding box it is contained within.
/// If this attribute is omitted, then a value of `overflow` is implied.
///
/// * Clip
/// * Overflow
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextHorizontalOverflowValues {
    Clip,
    Overflow,
}

impl TextHorizontalOverflowValues {
    pub(crate) fn default() -> Self {
        Self::Overflow
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "clip" => Self::Clip,
            "overflow" => Self::Overflow,
            _ => Self::default(),
        };
    }
}
