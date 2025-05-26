#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathshadevalues?view=openxml-3.0.1
///
/// * Circle (Radial)
/// * Rectangle (Rectangular)
/// * Shape (Shape Path)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PathShadeValues {
    Circle,
    Rectangle,
    Shape,
}

impl PathShadeValues {
    pub(crate) fn default() -> Self {
        Self::Shape
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "circle" => Self::Circle,
            "rect" => Self::Rectangle,
            "shape" => Self::Shape,
            _ => Self::default(),
        };
    }
}
