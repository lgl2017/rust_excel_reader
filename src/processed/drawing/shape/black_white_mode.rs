#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blackwhitemodevalues?view=openxml-3.0.1
///
/// * Auto
/// * Black
/// * BlackGray
/// * BlackWhite
/// * Color
/// * Gray
/// * GrayWhite
/// * Hidden
/// * InverseGray
/// * LightGray
/// * White
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum BlackWhiteModeValues {
    Auto,
    Black,
    BlackGray,
    BlackWhite,
    Color,
    Gray,
    GrayWhite,
    Hidden,
    InverseGray,
    LightGray,
    White,
}

impl BlackWhiteModeValues {
    pub(crate) fn default() -> Self {
        Self::Auto
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "auto" => Self::Auto,
            "black" => Self::Black,
            "blackGray" => Self::BlackGray,
            "blackWhite" => Self::BlackWhite,
            "clr" => Self::Color,
            "gray" => Self::Gray,
            "grayWhite" => Self::GrayWhite,
            "hidden" => Self::Hidden,
            "invGray" => Self::InverseGray,
            "ltGray" => Self::LightGray,
            "white" => Self::White,
            _ => Self::default(),
        };
    }
}
