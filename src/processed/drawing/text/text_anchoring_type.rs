#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textanchoringtypevalues?view=openxml-3.0.1
///
/// Specifies the anchoring position of the txBody within the shape.
/// If this attribute is omitted, then a value of t, meaning top, is implied.
///
/// * Bottom
/// * Center
/// * Top
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextAnchoringTypeValues {
    Bottom,
    Center,
    Top,
}

impl TextAnchoringTypeValues {
    pub(crate) fn default() -> Self {
        Self::Top
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "t" => Self::Top,
            "ctr" => Self::Center,
            "b" => Self::Bottom,
            _ => Self::default(),
        };
    }
}
