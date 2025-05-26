#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    common_types::HexColor,
    raw::{
        drawing::scheme::color_scheme::XlsxColorScheme,
        spreadsheet::{
            string_item::run_properties::XlsxRunProperties,
            stylesheet::{color::stylesheet_colors::XlsxStyleSheetColors, font::XlsxFont},
        },
    },
};

static DEFAULT_TEXT_COLOR: &str = "000000ff";
static DEFAULT_FONT_NAME: &str = "Calibri";
static DEFAULT_FONT_SIZE: f64 = 11.0;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-3.0.1
///
/// Example:
/// ```
/// <fonts count="2">
///     <font>
///         <sz val="11"/>
///         <color theme="1"/>
///         <name val="Calibri"/>
///         <family val="2"/>
///         <scheme val="minor"/>
///     </font>
///     <font>
///         <strike/>
///         <sz val="12"/>
///         <color theme="1"/>
///         <name val="Arial"/>
///         <family val="2"/>
///     </font>
/// </fonts>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Font {
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.bold?view=openxml-3.0.1
    pub bold: bool,

    ///  color
    pub color: HexColor,

    /// Condense: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.condense?view=openxml-3.0.1
    pub condense: bool,

    /// Extend: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.extend?view=openxml-3.0.1
    pub extend: bool,

    /// FontFamily: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontfamily?view=openxml-3.0.1
    pub family: FontFamilyValue,

    /// Italic: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.italic?view=openxml-3.0.1
    pub italic: bool,

    /// FontName: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontname?view=openxml-3.0.1
    pub name: String,

    /// Outline: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.outline?view=openxml-3.0.1
    pub outline: bool,

    /// FontScheme: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontscheme?view=openxml-3.0.1
    pub scheme: FontSchemeValue,

    /// Shadow: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.shadow?view=openxml-3.0.1
    pub shadow: bool,

    /// Strike: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.strike?view=openxml-3.0.1
    pub strike: bool,

    /// FontSize: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontsize?view=openxml-3.0.1
    pub size: f64,

    /// Underline: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.underline?view=openxml-3.0.1
    pub underline: UnderlineValue,

    /// VerticalTextAlignment: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.verticaltextalignment?view=openxml-3.0.1
    pub vertical_alignment: VerticalAlignmentRunValue,
}

impl Font {
    pub(crate) fn default() -> Self {
        Self {
            bold: false,
            color: DEFAULT_TEXT_COLOR.to_string(),
            condense: false,
            extend: false,
            family: FontFamilyValue::NotApplicable,
            italic: false,
            name: DEFAULT_FONT_NAME.to_string(),
            outline: false,
            scheme: FontSchemeValue::None,
            shadow: false,
            strike: false,
            size: DEFAULT_FONT_SIZE,
            underline: UnderlineValue::None,
            vertical_alignment: VerticalAlignmentRunValue::Baseline,
        }
    }

    pub(crate) fn from_raw_font(
        font: Option<XlsxFont>,
        stylesheet_colors: Option<XlsxStyleSheetColors>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        let Some(font) = font else {
            return Self::default();
        };
        let color_hex: Option<HexColor> = match font.color {
            Some(c) => c.to_hex(stylesheet_colors, color_scheme),
            None => None,
        };

        return Self {
            bold: font.bold.unwrap_or(false),
            color: color_hex.unwrap_or(DEFAULT_TEXT_COLOR.to_string()),
            condense: font.condense.unwrap_or(false),
            extend: font.extend.unwrap_or(false),
            family: FontFamilyValue::from_index(font.family),
            italic: font.italic.unwrap_or(false),
            name: font.name.unwrap_or(DEFAULT_FONT_NAME.to_string()),
            outline: font.outline.unwrap_or(false),
            scheme: FontSchemeValue::from_string(font.scheme),
            shadow: font.shadow.unwrap_or(false),
            strike: font.strike.unwrap_or(false),
            size: font.size.unwrap_or(DEFAULT_FONT_SIZE),
            underline: UnderlineValue::from_string(font.underline),
            vertical_alignment: VerticalAlignmentRunValue::from_string(font.vert_align),
        };
    }

    pub(crate) fn from_raw_run_properties(
        r_pr: Option<XlsxRunProperties>,
        stylesheet_colors: Option<XlsxStyleSheetColors>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        let Some(r_pr) = r_pr else {
            return Self::default();
        };
        let color_hex: Option<HexColor> = match r_pr.color {
            Some(c) => c.to_hex(stylesheet_colors, color_scheme),
            None => None,
        };

        return Self {
            bold: r_pr.bold.unwrap_or(false),
            color: color_hex.unwrap_or(DEFAULT_TEXT_COLOR.to_string()),
            condense: r_pr.condense.unwrap_or(false),
            extend: r_pr.extend.unwrap_or(false),
            family: FontFamilyValue::from_index(r_pr.family),
            italic: r_pr.italic.unwrap_or(false),
            name: r_pr.run_font.unwrap_or(DEFAULT_FONT_NAME.to_string()),
            outline: r_pr.outline.unwrap_or(false),
            scheme: FontSchemeValue::from_string(r_pr.scheme),
            shadow: r_pr.shadow.unwrap_or(false),
            strike: r_pr.strike.unwrap_or(false),
            size: r_pr.size.unwrap_or(DEFAULT_FONT_SIZE),
            underline: UnderlineValue::from_string(r_pr.underline),
            vertical_alignment: VerticalAlignmentRunValue::from_string(r_pr.vert_align),
        };
    }
}

/// FontFamily: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontfamily?view=openxml-3.0.1
///
/// Example:
/// ```
/// <family val="2"/>
/// ```
///
/// * 0: Not applicable.
/// * 1: Roman
/// * 2: Swiss
/// * 3: Modern
/// * 4: Script
/// * 5: Decorative
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum FontFamilyValue {
    NotApplicable,
    Roman,
    Swiss,
    Modern,
    Script,
    Decorative,
}

impl FontFamilyValue {
    pub(crate) fn from_index(index: Option<u64>) -> Self {
        let Some(index) = index else {
            return Self::NotApplicable;
        };
        return match index {
            0 => Self::NotApplicable,
            1 => Self::Roman,
            2 => Self::Swiss,
            3 => Self::Modern,
            4 => Self::Script,
            5 => Self::Decorative,
            _ => Self::NotApplicable,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.underlinevalues?view=openxml-3.0.1
///
/// * Double
/// * DoubleAccounting
/// * None
/// * Single
/// * SingleAccounting
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum UnderlineValue {
    Double,
    DoubleAccounting,
    None,
    Single,
    SingleAccounting,
}

impl UnderlineValue {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::None };
        return match s.as_ref() {
            "double" => Self::Double,
            "doubleAccounting" => Self::DoubleAccounting,
            "none" => Self::None,
            "single" => Self::Single,
            "singleAccounting" => Self::SingleAccounting,
            _ => Self::None,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontschemevalues?view=openxml-3.0.1
///
/// * Major
/// * Minor
/// * None
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum FontSchemeValue {
    Major,
    Minor,
    None,
}

impl FontSchemeValue {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::None };
        return match s.as_ref() {
            "major" => Self::Major,
            "minor" => Self::Minor,
            "none" => Self::None,
            _ => Self::None,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.verticalalignmentrunvalues?view=openxml-3.0.1
///
/// * Baseline,
/// * Subscript,
/// * Superscript
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum VerticalAlignmentRunValue {
    Baseline,
    Subscript,
    Superscript,
}

impl VerticalAlignmentRunValue {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::Baseline };
        return match s.as_ref() {
            "baseline" => Self::Baseline,
            "subscript" => Self::Subscript,
            "superscript" => Self::Superscript,
            _ => Self::Baseline,
        };
    }
}
