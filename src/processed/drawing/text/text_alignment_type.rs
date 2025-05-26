#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textalignmenttypevalues?view=openxml-3.0.1
///
/// * Center
/// * Distributed
/// * Justified
/// * JustifiedLow
/// * Left
/// * Right
/// * ThaiDistributed
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextAlignmentTypeValues {
    Center,
    Distributed,
    Justified,
    JustifiedLow,
    Left,
    Right,
    ThaiDistributed,
}

impl TextAlignmentTypeValues {
    pub(crate) fn default() -> Self {
        Self::Left
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "ctr" => Self::Center,
            "dist" => Self::Distributed,
            "just" => Self::Justified,
            "justLow" => Self::JustifiedLow,
            "l" => Self::Left,
            "r" => Self::Right,
            "thaiDist" => Self::ThaiDistributed,
            _ => Self::default(),
        };
    }
}
