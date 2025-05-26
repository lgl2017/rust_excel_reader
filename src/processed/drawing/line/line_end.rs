#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::line::head_end::XlsxHeadEnd;

/// * Tailend: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tailend?view=openxml-3.0.1
///
/// Example
/// ```
/// <tailEnd len="lg" type="arrowhead" w="sm"/>
/// ```
///
/// * Headend: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.headend?view=openxml-3.0.1
///
///  Example
/// ```
/// <headEnd len="lg" type="arrowhead" w="sm"/>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct LineEnd {
    /// Specifies the line end decoration, such as a triangle or arrowhead.
    pub r#type: LineEndTypeValue,

    /// Specifies the line end width in relation to the line width
    pub width: LineEndWidthValue,

    /// Specifies the line end length in relation to the line width.
    pub length: LineEndLengthValue,
}

impl LineEnd {
    pub(crate) fn default() -> Self {
        return Self {
            r#type: LineEndTypeValue::default(),
            width: LineEndWidthValue::default(),
            length: LineEndLengthValue::default(),
        };
    }

    pub(crate) fn from_raw(raw: Option<XlsxHeadEnd>, reference: Option<Self>) -> Self {
        let Some(raw) = raw else {
            return reference.unwrap_or(Self::default());
        };
        return Self {
            r#type: LineEndTypeValue::from_string(raw.clone().r#type),
            width: LineEndWidthValue::from_string(raw.clone().w),
            length: LineEndLengthValue::from_string(raw.clone().len),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineendvalues?view=openxml-3.0.1
///
/// * Arrow
/// * Diamond
/// * None
/// * Oval
/// * Stealth
/// * Triangle
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LineEndTypeValue {
    Arrow,
    Diamond,
    None,
    Oval,
    Stealth,
    Triangle,
}

impl LineEndTypeValue {
    pub(crate) fn default() -> Self {
        Self::None
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "arrow" => Self::Arrow,
            "diamond" => Self::Diamond,
            "none" => Self::None,
            "oval" => Self::Oval,
            "stealth" => Self::Stealth,
            "triangle" => Self::Triangle,
            _ => Self::default(),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineendlengthvalues?view=openxml-3.0.1
///
/// * Large
/// * Medium
/// * Small
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LineEndLengthValue {
    Large,
    Medium,
    Small,
}

impl LineEndLengthValue {
    pub(crate) fn default() -> Self {
        Self::Medium
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "lg" => Self::Large,
            "med" => Self::Medium,
            "sm" => Self::Small,
            _ => Self::default(),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineendwidthvalues?view=openxml-3.0.1
///
/// * Large
/// * Medium
/// * Small
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LineEndWidthValue {
    Large,
    Medium,
    Small,
}

impl LineEndWidthValue {
    pub(crate) fn default() -> Self {
        Self::Medium
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "lg" => Self::Large,
            "med" => Self::Medium,
            "sm" => Self::Small,
            _ => Self::default(),
        };
    }
}
