#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textunderlinevalues?view=openxml-3.0.1
///
/// * Dash
/// * DashHeavy
/// * DashLong
/// * DashLongHeavy
/// * DotDash
/// * DotDashHeavy
/// * DotDotDash
/// * DotDotDashHeavy
/// * Dotted
/// * Double
/// * Heavy
/// * HeavyDotted
/// * None
/// * Single
/// * Wavy
/// * WavyDouble
/// * WavyHeavy
/// * Words
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextUnderlineValues {
    Dash,
    DashHeavy,
    DashLong,
    DashLongHeavy,
    DotDash,
    DotDashHeavy,
    DotDotDash,
    DotDotDashHeavy,
    Dotted,
    Double,
    Heavy,
    HeavyDotted,
    None,
    Single,
    Wavy,
    WavyDouble,
    WavyHeavy,
    Words,
}

impl TextUnderlineValues {
    pub(crate) fn default() -> Self {
        Self::None
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "dash" => Self::Dash,
            "dashHeavy" => Self::DashHeavy,
            "dashLong" => Self::DashLong,
            "dashLongHeavy" => Self::DashLongHeavy,
            "dotDash" => Self::DotDash,
            "dotDashHeavy" => Self::DotDashHeavy,
            "dotDotDash" => Self::DotDotDash,
            "dotDotDashHeavy" => Self::DotDotDashHeavy,
            "dotted" => Self::Dotted,
            "dbl" => Self::Double,
            "heavy" => Self::Heavy,
            "dottedHeavy" => Self::HeavyDotted,
            "none" => Self::None,
            "sng" => Self::Single,
            "wavy" => Self::Wavy,
            "wavyDbl" => Self::WavyDouble,
            "wavyHeavy" => Self::WavyHeavy,
            "words" => Self::Words,
            _ => Self::default(),
        };
    }
}
