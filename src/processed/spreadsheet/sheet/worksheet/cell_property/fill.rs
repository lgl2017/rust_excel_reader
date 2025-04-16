use crate::{
    common_types::HexColor,
    raw::{
        drawing::scheme::color_scheme::ColorScheme,
        spreadsheet::stylesheet::{
            color::stylesheet_colors::StyleSheetColors,
            fill::{
                gradient_fill::{GradientFill as RawGradientFill, GradientStop as RawGradientStop},
                pattern_fill::PatternFill as RawPatternFill,
                Fill as RawFill,
            },
        },
    },
};

static DEFAULT_BACKGROUN_COLOR: &str = "ffffffff";

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fill?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub enum Fill {
    PatternFill(PatternFill),
    GradientFill(GradientFill),
}

impl Fill {
    pub(crate) fn default() -> Self {
        return Self::PatternFill(PatternFill {
            pattern_type: PatternFillTypeValue::None,
            foreground_color: None,
            background_color: None,
        });
    }
    pub(crate) fn from_raw(
        fill: Option<RawFill>,
        stylesheet_colors: Option<StyleSheetColors>,
        color_scheme: Option<ColorScheme>,
    ) -> Self {
        let Some(fill) = fill else {
            return Self::default();
        };
        return match fill {
            RawFill::PatternFill(pattern_fill) => Self::PatternFill(PatternFill::from_raw(
                pattern_fill,
                stylesheet_colors,
                color_scheme,
            )),
            RawFill::GradientFill(gradient_fill) => Self::GradientFill(GradientFill::from_raw(
                gradient_fill,
                stylesheet_colors,
                color_scheme,
            )),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <fill>
///     <patternFill patternType="solid">
///         <fgColor indexed="12" />
///         <bgColor auto="1" />
///     </patternFill>
///// </fill>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct PatternFill {
    /// Specifies the fill pattern type (including solid and none).
    /// Default is none, when missing
    pub pattern_type: PatternFillTypeValue,

    /// foreground color
    pub foreground_color: Option<HexColor>,

    /// background color
    pub background_color: Option<HexColor>,
}

impl PatternFill {
    pub(crate) fn from_raw(
        fill: RawPatternFill,
        stylesheet_colors: Option<StyleSheetColors>,
        color_scheme: Option<ColorScheme>,
    ) -> Self {
        let foreground_color: Option<HexColor> = match fill.foreground_color {
            Some(c) => c.to_hex(stylesheet_colors.clone(), color_scheme.clone()),
            None => None,
        };
        let mut background_color: Option<HexColor> = match fill.background_color {
            Some(c) => c.to_hex(stylesheet_colors.clone(), color_scheme.clone()),
            None => None,
        };

        let pattern_type = PatternFillTypeValue::from_string(fill.pattern_type);
        if pattern_type != PatternFillTypeValue::None && background_color.is_none() {
            background_color = Some(DEFAULT_BACKGROUN_COLOR.to_string())
        };

        return Self {
            pattern_type,
            foreground_color,
            background_color,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternvalues?view=openxml-3.0.1
///
/// * DarkDown
/// * DarkGray
/// * DarkGrid
/// * DarkHorizontal
/// * DarkTrellis
/// * DarkUp
/// * DarkVertical
/// * Gray0625
/// * Gray125
/// * LightDown
/// * LightGray
/// * LightGrid
/// * LightHorizontal
/// * LightTrellis
/// * LightUp
/// * LightVertical
/// * MediumGray
/// * None
/// * Solid
#[derive(Debug, Clone, PartialEq)]
pub enum PatternFillTypeValue {
    DarkDown,
    DarkGray,
    DarkGrid,
    DarkHorizontal,
    DarkTrellis,
    DarkUp,
    DarkVertical,
    Gray0625,
    Gray125,
    LightDown,
    LightGray,
    LightGrid,
    LightHorizontal,
    LightTrellis,
    LightUp,
    LightVertical,
    MediumGray,
    None,
    Solid,
}

impl PatternFillTypeValue {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::None };
        return match s.as_ref() {
            "darkDown" => Self::DarkDown,
            "darkGray" => Self::DarkGray,
            "darkGrid" => Self::DarkGrid,
            "darkHorizontal" => Self::DarkHorizontal,
            "darkTrellis" => Self::DarkTrellis,
            "darkUp" => Self::DarkUp,
            "darkVertical" => Self::DarkVertical,
            "gray0625" => Self::Gray0625,
            "gray125" => Self::Gray125,
            "lightDown" => Self::LightDown,
            "lightGray" => Self::LightGray,
            "lightGrid" => Self::LightGrid,
            "lightHorizontal" => Self::LightHorizontal,
            "lightTrellis" => Self::LightTrellis,
            "lightUp" => Self::LightUp,
            "lightVertical" => Self::LightVertical,
            "mediumGray" => Self::MediumGray,
            "none" => Self::None,
            "solid" => Self::Solid,
            _ => Self::None,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.gradientfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <fill>
///     <gradientFill degree="90">
///         <stop position="0">
///             <color rgb="FF92D050"/>
///         </stop>
///         <stop position="1">
///             <color rgb="FF0070C0"/>
///         </stop>
///     </gradientFill>
/// </fill>
/// <fill>
///     <gradientFill type="path" left="0.2" right="0.8" top="0.2" bottom="0.8">
///         <stop position="0">
///             <color theme="0"/>
///         </stop>
///         <stop position="1">
///             <color theme="4"/>
///         </stop>
///     </gradientFill>
/// </fill>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct GradientFill {
    /// Specifies the position of the bottom edge of the inner rectangle
    /// For bottom, 0 means the bottom edge of the inner rectangle is on the top edge of the cell, and 1 means it is on the bottom edge of the cell. (applies to From Corner and From Center gradients).
    ///
    /// values ranging from 0 to 1.
    pub bottom: Option<f64>,

    /// Specifies the position of the left edge of the inner rectangle
    /// For left, 0 means the left edge of the inner rectangle is on the left edge of the cell, and 1 means it is on the right edge of the cell. (applies to From Corner and From Center gradients).
    ///
    /// values ranging from 0 to 1.
    pub left: Option<f64>,

    /// Specifies the position of the right edge of the inner rectangle
    /// For right, 0 means the right edge of the inner rectangle is on the left edge of the cell, and 1 means it is on the right edge of the cell. (applies to From Corner and From Center gradients).
    ///
    /// values ranging from 0 to 1.
    pub right: Option<f64>,

    /// Specifies the position of the top edge of the inner rectangle
    /// For top, 0 means the top edge of the inner rectangle is on the top edge of the cell, and 1 means it is on the bottom edge of the cell. (applies to From Corner and From Center gradients).
    ///
    /// values ranging from 0 to 1.
    pub top: Option<f64>,

    /// Angle of the linear gradient
    pub degree: Option<f64>,

    /// Type of gradient fill.
    pub r#type: GradientFillTypeValue,

    /// * children
    pub stop: Option<Vec<GradientStop>>,
}

impl GradientFill {
    pub(crate) fn from_raw(
        fill: RawGradientFill,
        stylesheet_colors: Option<StyleSheetColors>,
        color_scheme: Option<ColorScheme>,
    ) -> Self {
        let stops: Option<Vec<GradientStop>> = if let Some(stops) = fill.stop {
            let proccessed: Vec<GradientStop> = stops
                .into_iter()
                .map(|s| GradientStop::from_raw(s, stylesheet_colors.clone(), color_scheme.clone()))
                .collect();
            if proccessed.is_empty() {
                None
            } else {
                Some(proccessed)
            }
        } else {
            None
        };
        return Self {
            bottom: fill.bottom,
            left: fill.left,
            right: fill.right,
            top: fill.top,
            degree: fill.degree,
            r#type: GradientFillTypeValue::from_string(fill.r#type),
            stop: stops,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.gradientvalues?view=openxml-3.0.1
///
/// * Linear
/// * Path
#[derive(Debug, Clone, PartialEq)]
pub enum GradientFillTypeValue {
    Linear,
    Path,
}

impl GradientFillTypeValue {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::Linear };
        return match s.as_ref() {
            "linear" => Self::Linear,
            "path" => Self::Path,
            _ => Self::Linear,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.gradientstop?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct GradientStop {
    /// Position information for this gradient stop
    ///
    /// Interpreted exactly like gradientFill left, right, bottom, top.
    /// The position indicated here indicates the point where the color is pure.
    /// Before and and after this position the color can be in transition (or pure, depending on if this is the last stop or not).
    pub position: Option<f64>,

    pub color: Option<HexColor>,
}

impl GradientStop {
    pub(crate) fn from_raw(
        stop: RawGradientStop,
        stylesheet_colors: Option<StyleSheetColors>,
        color_scheme: Option<ColorScheme>,
    ) -> Self {
        let color_hex: Option<HexColor> = match stop.color {
            Some(c) => c.to_hex(stylesheet_colors, color_scheme),
            None => None,
        };

        return Self {
            position: stop.position,
            color: color_hex,
        };
    }
}
