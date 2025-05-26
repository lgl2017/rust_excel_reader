#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textfontalignmentvalues?view=openxml-3.0.1
///
/// * Automatic
/// * Baseline
/// * Bottom
/// * Center
/// * Top
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextFontAlignmentValues {
    Automatic,
    Baseline,
    Bottom,
    Center,
    Top,
}

impl TextFontAlignmentValues {
    pub(crate) fn default() -> Self {
        Self::Automatic
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "auto" => Self::Automatic,
            "base" => Self::Baseline,
            "b" => Self::Bottom,
            "ctr" => Self::Center,
            "t" => Self::Top,
            _ => Self::default(),
        };
    }
}
