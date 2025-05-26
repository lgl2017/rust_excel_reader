#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathfillmodevalues?view=openxml-3.0.1
///
/// Specifies how the path should be filled.
/// Default to "norm".
///
/// * Darken: Darken Path Fill.
/// * DarkenLess: Darken Path Fill Less.
/// * Lighten: Lighten Path Fill.
/// * LightenLess: Lighten Path Fill Less.
/// * None: No Path Fill.
/// * Norm: Normal Path Fill.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PathFillModeValues {
    Darken,
    DarkenLess,
    Lighten,
    LightenLess,
    None,
    Norm,
}

impl PathFillModeValues {
    pub(crate) fn default() -> Self {
        Self::Norm
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "darken" => Self::Darken,
            "darkenLess" => Self::DarkenLess,
            "lighten" => Self::Lighten,
            "lightenLess" => Self::LightenLess,
            "none" => Self::None,
            "norm" => Self::Norm,
            _ => Self::default(),
        };
    }
}
