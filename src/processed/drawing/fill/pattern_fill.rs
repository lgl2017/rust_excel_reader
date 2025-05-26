#[cfg(feature = "serde")]
use serde::Serialize;

use crate::common_types::HexColor;
use crate::raw::drawing::fill::pattern_fill::XlsxPatternFill;
use crate::raw::drawing::scheme::color_scheme::XlsxColorScheme;

use super::preset_pattern_values::PresetPatternValues;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.patternfill?view=openxml-3.0.1
///
/// Specifies a pattern fill.
/// A repeated pattern is used to fill the object.
///
/// Example:
/// ```
/// <a:pattFill prst="cross">
///     <a:bgClr>
///         <a:solidFill>
///             <a:schemeClr val="phClr" />
///         </a:solidFill>
///     </a:bgClr>
///     <a:fgClr>
///         <a:schemeClr val="phClr">
///             <a:satMod val="110000" />
///             <a:lumMod val="100000" />
///             <a:shade val="100000" />
///         </a:schemeClr>
///     </a:fgClr>
/// </a:pattFill>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct PatternFill {
    /// ForegroundColor
    pub foreground_color: HexColor,

    /// BackgroundColor
    pub background_color: HexColor,

    /// Specifies one of a set of preset patterns to fill the object
    pub pattern: PresetPatternValues,
}

impl PatternFill {
    pub(crate) fn from_raw(
        raw: Option<XlsxPatternFill>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };

        let Some(pattern) = PresetPatternValues::from_string(raw.prst) else {
            return None;
        };
        let fg = if let Some(fg) = raw.fg_clr {
            fg.to_hex(color_scheme.clone(), ref_color.clone())
        } else if let Some(color_scheme) = color_scheme.clone() {
            color_scheme.accent5
        } else {
            None
        };
        let bg = if let Some(bg) = raw.bg_clr {
            bg.to_hex(color_scheme.clone(), ref_color.clone())
        } else {
            None
        };

        return Some(Self {
            foreground_color: bg.unwrap_or("ffffffff".to_string()), // bg1
            background_color: fg.unwrap_or("a02b93ff".to_string()), // accent5
            pattern,
        });
    }
}
