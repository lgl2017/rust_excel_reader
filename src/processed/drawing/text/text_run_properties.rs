#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    processed::{
        drawing::{effect::effect_container::EffectContainer, fill::Fill, line::outline::Outline},
        shared::hyperlink::Hyperlink,
    },
    raw::{
        drawing::{
            scheme::color_scheme::XlsxColorScheme,
            st_types::{st_percentage_to_float, st_text_point_to_pt},
            text::default_text_run_properties::XlsxTextRunProperties,
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::{
    font::Font, text_cap_values::TextCapsValues, text_strike_values::TextStrikeValues,
    underline::Underline,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.defaultrunproperties?view=openxml-3.0.1
///
/// This element contains all default run level text properties for the text runs within a containing paragraph.
/// These properties are to be used when overriding properties have not been defined within the rPr element
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct TextRunProperties {
    /// font typeface
    pub font_typeface: Font,

    /// FontSize (pts)
    pub font_size: f64,

    /// Language
    pub language: String,

    /// fill of text
    pub fill: Fill,

    /// underline of text run
    pub underline: Underline,

    /// Strike
    pub strike: TextStrikeValues,

    /// Bold
    pub bold: bool,

    /// Italic
    pub italic: bool,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.righttoleft?view=openxml-3.0.1
    ///
    /// This element specifies whether the contents of this run shall have right-to-left characteristics.
    pub right_to_left: bool,

    /// Baseline (percentage)
    pub baseline: f64,

    /// Capital
    pub capitalize: TextCapsValues,

    /// Spacing (pts)
    pub spacing: f64,

    /// outline on text
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub outline: Option<Outline>,

    /// effect apply to the text run
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub effect_list: Option<Box<EffectContainer>>,

    /// hyperlink on drawing click
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub hyperlink_on_click: Option<Hyperlink>,

    /// hyperlink on drawing hover
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub hyperlink_on_hover: Option<Hyperlink>,

    /// highlight
    ///
    /// This element specifies the highlight color that is present for a run of text.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.highlight?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub highlight_color: Option<HexColor>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.symbolfont?view=openxml-3.0.1
    ///
    /// This element specifies that a symbol font typeface be used for a specific run of text.
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub symbol_font: Option<String>,

    /// AlternativeLanguage
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub alternative_language: Option<String>,

    /// Bookmark
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub bookmark: Option<String>,

    /// Dirty
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub dirty: Option<bool>,

    /// Kerning
    ///
    /// Use kerning for font above the given size
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub kerning_font_size: Option<f64>,

    /// Kumimoji
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub kumimoji: Option<bool>,

    /// NoProof
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub no_proof: Option<bool>,

    /// NormalizeHeight
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub normalize_height: Option<bool>,

    /// SmartTagClean
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub smart_tag_clean: Option<bool>,

    // SmartTagId
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub smart_tag_id: Option<u64>,

    /// SpellingError
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub spelling_error: Option<bool>,
}

impl TextRunProperties {
    pub(crate) fn default() -> Self {
        return Self {
            font_typeface: Font::default(),
            font_size: 11.0,
            language: "en-US".to_string(),
            fill: Fill::SolidFill("000000ff".to_string()),
            underline: Underline::default(),
            strike: TextStrikeValues::default(),
            bold: false,
            italic: false,
            right_to_left: false,
            baseline: 0.0,
            capitalize: TextCapsValues::default(),
            spacing: 0.0,
            outline: None,
            effect_list: None,
            hyperlink_on_click: None,
            hyperlink_on_hover: None,
            highlight_color: None,
            symbol_font: None,
            alternative_language: None,
            bookmark: None,
            dirty: None,
            kerning_font_size: None,
            kumimoji: None,
            no_proof: None,
            normalize_height: None,
            smart_tag_clean: None,
            smart_tag_id: None,
            spelling_error: None,
        };
    }

    pub(crate) fn from_raw(
        raw: Option<XlsxTextRunProperties>,
        default_properties: Option<Self>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_font: Option<Font>,
        font_ref_color: Option<HexColor>,
    ) -> Self {
        let default = if let Some(default) = default_properties {
            default
        } else {
            Self::default()
        };

        let Some(raw) = raw else {
            return default;
        };

        let fill = Fill::from_text_run_properties(
            raw.clone(),
            Some(default.clone()),
            drawing_relationship.clone(),
            image_bytes.clone(),
            color_scheme.clone(),
            font_ref_color.clone(),
        );

        let outline =
            if let Some(ln) = Outline::from_raw(raw.clone().outline, color_scheme.clone(), None) {
                Some(ln)
            } else {
                default.clone().outline
            };

        let right_to_left = if let Some(rtl) = raw.clone().rtl {
            rtl.val.unwrap_or(default.clone().right_to_left)
        } else {
            default.clone().right_to_left
        };

        return Self {
            font_typeface: Font::from_text_run_properties(
                raw.clone(),
                Some(default.clone()),
                ref_font.clone(),
            ),
            font_size: if let Some(st) = raw.clone().font_size {
                st_text_point_to_pt(st as i64)
            } else {
                default.clone().font_size
            },
            language: raw.clone().language.unwrap_or(default.clone().language),
            fill: fill.clone(),
            underline: Underline::from_text_run_properties(
                raw.clone(),
                Some(default.clone().underline),
                outline.clone(),
                fill.clone(),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                font_ref_color.clone(),
            ),
            strike: if let Some(s) = raw.clone().strike {
                TextStrikeValues::from_string(Some(s))
            } else {
                default.clone().strike
            },
            bold: raw.clone().bold.unwrap_or(default.clone().bold),
            italic: raw.clone().italic.unwrap_or(default.clone().italic),
            right_to_left,
            baseline: if let Some(st) = raw.clone().baseline {
                st_percentage_to_float(st as i64)
            } else {
                default.clone().baseline
            },
            capitalize: if let Some(s) = raw.clone().capital {
                TextCapsValues::from_string(Some(s))
            } else {
                default.clone().capitalize
            },
            spacing: if let Some(st) = raw.clone().spacing {
                st_text_point_to_pt(st as i64)
            } else {
                default.clone().font_size
            },
            outline: outline.clone(),
            effect_list: if let Some(e) = EffectContainer::from_raw_effect_list(
                raw.clone().effect_list,
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                None,
            ) {
                Some(e)
            } else {
                default.clone().effect_list
            },
            hyperlink_on_click: if let Some(h) = Hyperlink::from_hlink_event(
                raw.clone().hlink_click,
                drawing_relationship.clone(),
                defined_names.clone(),
            ) {
                Some(h)
            } else {
                default.clone().hyperlink_on_click
            },
            hyperlink_on_hover: if let Some(h) = Hyperlink::from_hlink_event(
                raw.clone().hlink_mouse_over,
                drawing_relationship.clone(),
                defined_names.clone(),
            ) {
                Some(h)
            } else {
                default.clone().hyperlink_on_click
            },
            highlight_color: if let Some(higlight) = raw.clone().highlight {
                higlight.to_hex(color_scheme.clone(), None)
            } else {
                default.clone().highlight_color
            },
            symbol_font: if let Some(f) = raw.clone().symbol_font {
                f.typeface
            } else {
                default.clone().symbol_font
            },
            alternative_language: if let Some(s) = raw.clone().alternative_language {
                Some(s)
            } else {
                default.clone().alternative_language
            },
            bookmark: if let Some(s) = raw.clone().bookmark {
                Some(s)
            } else {
                default.clone().bookmark
            },
            dirty: if let Some(b) = raw.clone().dirty {
                Some(b)
            } else {
                default.clone().dirty
            },
            kerning_font_size: if let Some(k) = raw.clone().kerning {
                Some(st_text_point_to_pt(k as i64))
            } else {
                default.clone().kerning_font_size
            },
            kumimoji: if let Some(b) = raw.clone().kumimoji {
                Some(b)
            } else {
                default.clone().kumimoji
            },
            no_proof: if let Some(b) = raw.clone().no_proof {
                Some(b)
            } else {
                default.clone().no_proof
            },
            normalize_height: if let Some(b) = raw.clone().normalize_height {
                Some(b)
            } else {
                default.clone().normalize_height
            },
            smart_tag_clean: if let Some(b) = raw.clone().smart_tag_clean {
                Some(b)
            } else {
                default.clone().smart_tag_clean
            },
            smart_tag_id: if let Some(val) = raw.clone().smart_tag_id {
                Some(val)
            } else {
                default.clone().smart_tag_id
            },
            spelling_error: if let Some(b) = raw.clone().spelling_error {
                Some(b)
            } else {
                default.clone().spelling_error
            },
        };
    }
}
