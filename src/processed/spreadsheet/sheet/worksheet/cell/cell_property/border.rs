use crate::{
    common_types::HexColor,
    raw::{
        drawing::scheme::color_scheme::XlsxColorScheme,
        spreadsheet::stylesheet::{
            border::{Border as RawBorder, BorderStyle as RawBorderStyle},
            color::stylesheet_colors::XlsxStyleSheetColors,
        },
    },
};

static DEFAULT_BORDER_COLOR: &str = "000000ff";

/// Border: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-3.0.1
///
/// Expresses a single set of cell border formats (left, right, top, bottom, diagonal).
/// Color is optional. When missing, 'automatic' is implied.
///
/// Example
/// ```
/// <border>
///     <left/>
///     <right/>
///     <top/>
///     <bottom/>
///     <diagonal/>
/// </border>
/// <border>
///     <left/>
///     <right style="medium">
///         <color indexed="64"/>
///     </right>
///     <top/>
///     <bottom style="thin">
///         <color indexed="64"/>
///     </bottom>
///     <diagonal/>
///     </border>
/// </borders>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Border {
    pub left: BorderStyle,

    pub right: BorderStyle,

    pub top: BorderStyle,

    pub bottom: BorderStyle,

    /// diagonal line starting at the top left corner of the cell and moving down to the bottom right corner of the cell.
    pub diagonal_down: BorderStyle,

    /// diagonal line starting at the bottom left corner of the cell and moving up to the top right corner of the cell.
    pub diagonal_up: BorderStyle,

    /// A boolean value indicating if left, right, top, and bottom borders should be applied only to outside borders of a cell range.
    pub outline: bool,
}

impl Border {
    pub(crate) fn default() -> Self {
        return Self {
            left: BorderStyle::default(),
            right: BorderStyle::default(),
            top: BorderStyle::default(),
            bottom: BorderStyle::default(),
            diagonal_down: BorderStyle::default(),
            diagonal_up: BorderStyle::default(),
            outline: false,
        };
    }

    pub(crate) fn from_raw(
        border: Option<RawBorder>,
        stylesheet_colors: Option<XlsxStyleSheetColors>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        let Some(border) = border else {
            return Self::default();
        };
        let diagonal = BorderStyle::from_raw(
            border.clone().diagonal,
            stylesheet_colors.clone(),
            color_scheme.clone(),
        );
        let diagonal_down = if border.clone().diagonal_down.unwrap_or(false) == true {
            diagonal.clone()
        } else {
            BorderStyle::default()
        };
        let diagonal_up = if border.clone().diagonal_up.unwrap_or(false) == true {
            diagonal.clone()
        } else {
            BorderStyle::default()
        };

        return Self {
            left: BorderStyle::from_raw(
                border.clone().left,
                stylesheet_colors.clone(),
                color_scheme.clone(),
            ),
            right: BorderStyle::from_raw(
                border.clone().right,
                stylesheet_colors.clone(),
                color_scheme.clone(),
            ),
            top: BorderStyle::from_raw(
                border.clone().top,
                stylesheet_colors.clone(),
                color_scheme.clone(),
            ),
            bottom: BorderStyle::from_raw(
                border.clone().bottom,
                stylesheet_colors.clone(),
                color_scheme.clone(),
            ),
            diagonal_down,
            diagonal_up,
            outline: border.outline.unwrap_or(false),
        };
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct BorderStyle {
    /// The line style for this border
    pub style: BorderStyleValue,
    pub color: Option<HexColor>,
}

impl BorderStyle {
    pub(crate) fn default() -> Self {
        return Self {
            style: BorderStyleValue::None,
            color: None,
        };
    }

    pub(crate) fn from_raw(
        style: Option<RawBorderStyle>,
        stylesheet_colors: Option<XlsxStyleSheetColors>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        let Some(style) = style else {
            return Self::default();
        };
        let mut color: Option<HexColor> = match style.color {
            Some(c) => c.to_hex(stylesheet_colors, color_scheme),
            None => None,
        };
        let border_style = BorderStyleValue::from_string(style.style);
        if border_style != BorderStyleValue::None && color.is_none() {
            color = Some(DEFAULT_BORDER_COLOR.to_owned())
        }

        return Self {
            style: border_style,
            color,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.borderstylevalues?view=openxml-3.0.1
///
/// * DashDot
/// * DashDotDot
/// * Dashed
/// * Dotted
/// * Double
/// * Hair
/// * Medium
/// * MediumDashDot
/// * MediumDashDotDot
/// * MediumDashed
/// * None
/// * SlantDashDot
/// * Thick
/// * Thin
#[derive(Debug, Clone, PartialEq)]
pub enum BorderStyleValue {
    DashDot,
    DashDotDot,
    Dashed,
    Dotted,
    Double,
    Hair,
    Medium,
    MediumDashDot,
    MediumDashDotDot,
    MediumDashed,
    None,
    SlantDashDot,
    Thick,
    Thin,
}

impl BorderStyleValue {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::None };
        return match s.as_ref() {
            "dashDot" => Self::DashDot,
            "dashDotDot" => Self::DashDotDot,
            "dashed" => Self::Dashed,
            "dotted" => Self::Dotted,
            "double" => Self::Double,
            "hair" => Self::Hair,
            "medium" => Self::Medium,
            "mediumDashDot" => Self::MediumDashDot,
            "mediumDashDotDot" => Self::MediumDashDotDot,
            "mediumDashed" => Self::MediumDashed,
            "none" => Self::None,
            "slantDashDot" => Self::SlantDashDot,
            "thick" => Self::Thick,
            "thin" => Self::Thin,
            _ => Self::None,
        };
    }
}
