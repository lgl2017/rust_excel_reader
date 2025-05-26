#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textstrikevalues?view=openxml-3.0.1
///
/// * DoubleStrike
/// * NoStrike
/// * SingleStrike
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextStrikeValues {
    DoubleStrike,
    NoStrike,
    SingleStrike,
}

impl TextStrikeValues {
    pub(crate) fn default() -> Self {
        Self::NoStrike
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "dblStrike" => Self::DoubleStrike,
            "noStrike" => Self::NoStrike,
            "sngStrike" => Self::SingleStrike,
            _ => Self::default(),
        };
    }
}
