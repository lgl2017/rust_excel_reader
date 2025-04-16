use std::usize;

use anyhow::bail;
use quick_xml::events::BytesStart;
use stylesheet_colors::{get_default_indexed_color_mapping, StyleSheetColors};

use crate::{
    common_types::HexColor,
    helper::{
        apply_tint, format_hex_string, hex_to_rgba, rgba_to_hex, string_to_bool, string_to_float,
        string_to_unsignedint,
    },
    raw::drawing::scheme::color_scheme::ColorScheme,
};

pub mod rgb_color;
pub mod stylesheet_colors;

/// BackgroundColor: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.backgroundcolor?view=openxml-3.0.1
pub type BackgroundColor = Color;

/// ForegroundColor: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.foregroundcolor?view=openxml-3.0.1
pub type ForegroundColor = Color;

/// Color corresponding to the following classes
///
/// DataBarColor: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.color?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct Color {
    // attributes
    /// A boolean value (0: false, 1: true) indicating the color is automatic and system color dependent.
    /// Example:
    /// ```
    /// <bgColor auto="1" />
    /// ```
    pub auto: Option<bool>,

    /// Indexed color value.
    pub indexed: Option<u64>,

    /// Standard Alpha Red Green Blue color value (ARGB).
    /// Ex: ffff95ca
    pub rgb: Option<String>,

    /// A zero-based index into the <clrScheme> collection (§20.1.6.2),
    /// referencing a particular <sysClr> or <srgbClr> value expressed in the Theme part.
    pub theme: Option<u64>,

    /// Specifies the tint value applied to the color
    ///
    /// If tint is supplied, then it is applied to the value of the color to determine the final color applied.
    /// The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and 1.0 means 100% lighten. Also, 0.0 means no change.
    pub tint: Option<f64>,
}

impl Color {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut color = Self {
            auto: None,
            indexed: None,
            rgb: None,
            theme: None,
            tint: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"auto" => color.auto = string_to_bool(&string_value),
                        b"indexed" => color.indexed = string_to_unsignedint(&string_value),
                        b"rgb" => color.rgb = Some(string_value),
                        b"theme" => color.theme = string_to_unsignedint(&string_value),
                        b"tint" => color.tint = string_to_float(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(color)
    }
}

impl Color {
    pub(crate) fn to_hex(
        &self,
        stylesheet_colors: Option<StyleSheetColors>,
        color_scheme: Option<ColorScheme>,
    ) -> Option<HexColor> {
        let Some(base_color) = self.get_base_color(stylesheet_colors, color_scheme) else {
            return None;
        };
        let Some(tint) = self.tint.clone() else {
            return Some(base_color);
        };
        let Ok(rgba) = hex_to_rgba(&base_color, Some(false)) else {
            return Some(base_color);
        };
        let Ok(rgba) = apply_tint(rgba, tint) else {
            return Some(base_color);
        };

        let Ok(hex) = rgba_to_hex(rgba, Some(false)) else {
            return Some(base_color);
        };

        return Some(hex);
    }

    fn get_base_color(
        &self,
        stylesheet_colors: Option<StyleSheetColors>,
        color_scheme: Option<ColorScheme>,
    ) -> Option<HexColor> {
        if let (Some(theme_index), Some(color_scheme)) = (self.theme, color_scheme) {
            return color_scheme.get_color(theme_index);
        }

        // <color rgb="ffff95ca" />
        if let Some(hex) = self.rgb.clone() {
            if let Ok(new) = format_hex_string(&hex, Some(true)) {
                return Some(new);
            };
        }

        if let (Some(index), Some(stylesheet_colors)) = (self.indexed, stylesheet_colors) {
            if let Some(hex) = stylesheet_colors.get_indexed_color(index) {
                return Some(hex);
            }
        };

        if let Some(index) = self.indexed {
            let default_mapping = get_default_indexed_color_mapping();
            if let Ok(index) = TryInto::<usize>::try_into(index) {
                if index < default_mapping.len() {
                    return Some(default_mapping[index].clone());
                }
            }
        };

        return None;
    }
}
