#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linecapvalues?view=openxml-3.0.1
///
/// * Flat
/// * Round
/// * Square
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LineCapValues {
    Flat,
    Round,
    Square,
}

impl LineCapValues {
    pub(crate) fn default() -> Self {
        Self::Square
    }

    pub(crate) fn from_string(s: Option<String>, reference: Option<Self>) -> Self {
        let Some(s) = s else {
            return reference.unwrap_or(Self::default());
        };
        return match s.as_ref() {
            "flat" => Self::Flat,
            "rnd" => Self::Round,
            "sq" => Self::Square,
            _ => Self::default(),
        };
    }
}
