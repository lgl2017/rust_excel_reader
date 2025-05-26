#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textwrappingvalues?view=openxml-3.0.1
///
/// Specifies the wrapping options to be used for this text body.
/// If this attribute is omitted, then a value of `square` is implied which will wrap the text using the bounding text box.
///
/// * None
/// * Square
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextWrappingValues {
    None,
    Square,
}

impl TextWrappingValues {
    pub(crate) fn default() -> Self {
        Self::Square
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "none" => Self::None,
            "square" => Self::Square,
            _ => Self::default(),
        };
    }
}
