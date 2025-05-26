#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textverticalvalues?view=openxml-3.0.1
///
/// Determines if the text within the given text body should be displayed vertically.
/// If this attribute is omitted, then a value of `horz`, meaning no vertical text, is implied.
///
/// * EastAsianVetical
/// * Horizontal
/// * MongolianVertical
/// * Vertical
/// * Vertical270
/// * WordArtLeftToRight
/// * WordArtVertical
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextVerticalValues {
    EastAsianVetical,
    Horizontal,
    MongolianVertical,
    Vertical,
    Vertical270,
    WordArtLeftToRight,
    WordArtVertical,
}

impl TextVerticalValues {
    pub(crate) fn default() -> Self {
        Self::Horizontal
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "eaVert" => Self::EastAsianVetical,
            "horz" => Self::Horizontal,
            "mongolianVert" => Self::MongolianVertical,
            "vert" => Self::Vertical,
            "vert270" => Self::Vertical270,
            "wordArtVertRtl" => Self::WordArtLeftToRight,
            "wordArtVert" => Self::WordArtVertical,
            _ => Self::default(),
        };
    }
}
