#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textcapsvalues?view=openxml-3.0.1
///
/// * All
/// * None
/// * Small
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextCapsValues {
    All,
    None,
    Small,
}

impl TextCapsValues {
    pub(crate) fn default() -> Self {
        Self::None
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "all" => Self::All,
            "none" => Self::None,
            "small" => Self::Small,
            _ => Self::default(),
        };
    }
}
