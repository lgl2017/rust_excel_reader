#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    common_types::HexColor,
    raw::drawing::{effect::color_change::XlsxColorChange, scheme::color_scheme::XlsxColorScheme},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorchange?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:clrChange useA="1">
///     <a:clrFrom>
///         <a:schemeClr val="phClr" />
///     </a:clrFrom>
///     <a:clrTo>
///         <a:schemeClr val="phClr" />
///     </a:clrTo>
/// </a:clrChange>
/// ```
// tag: clrChange
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ColorChange {
    // Change Color From
    pub color_from: HexColor,

    // Change Color To
    pub color_to: HexColor,

    /// Specifies whether alpha values are considered for the effect.
    /// Effect alpha values are considered if useA is true, else they are ignored.
    pub use_alpha: bool,
}

impl ColorChange {
    pub(crate) fn from_raw(
        raw: Option<XlsxColorChange>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };

        let (Some(to), Some(from)) = (raw.color_to, raw.color_from) else {
            return None;
        };

        let (Some(hex_to), Some(hex_from)) = (
            to.to_hex(color_scheme.clone(), ref_color.clone()),
            from.to_hex(color_scheme.clone(), ref_color.clone()),
        ) else {
            return None;
        };

        return Some(Self {
            color_from: hex_from,
            color_to: hex_to,
            use_alpha: raw.use_a.unwrap_or(false),
        });
    }
}
