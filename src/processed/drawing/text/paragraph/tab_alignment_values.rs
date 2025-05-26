#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.texttabalignmentvalues?view=openxml-3.0.1
///
/// * Center
/// * Decimal
/// * Left
/// * Right
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextTabAlignmentValues {
    Center,
    Decimal,
    Left,
    Right,
}

impl TextTabAlignmentValues {
    pub(crate) fn default() -> Self {
        Self::Left
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "ctr" => Self::Center,
            "dec" => Self::Decimal,
            "l" => Self::Left,
            "r" => Self::Right,
            _ => Self::default(),
        };
    }
}
