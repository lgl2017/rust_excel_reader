#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tileflipvalues?view=openxml-3.0.1
///
/// * Horizontal
/// * HorizontalAndVertical
/// * None
/// * Vertical
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TileFlipValues {
    Horizontal,
    HorizontalAndVertical,
    None,
    Vertical,
}

impl TileFlipValues {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::None };
        return match s.as_ref() {
            "x" => Self::Horizontal,
            "xy" => Self::HorizontalAndVertical,
            "none" => Self::None,
            "y" => Self::Vertical,
            _ => Self::None,
        };
    }
}
