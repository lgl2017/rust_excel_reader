#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    packaging::relationship::XlsxRelationships,
    processed::drawing::text::{
        font_alignment_values::TextFontAlignmentValues,
        text_alignment_type::TextAlignmentTypeValues,
    },
    raw::drawing::{
        scheme::color_scheme::XlsxColorScheme, st_types::emu_to_pt,
        text::paragraph::paragraph_properties::XlsxParagraphProperties,
    },
};

use super::{bullet::Bullet, spacing_type::SpacingTypeValues, tab_stop::TabStop};

// There are a total of 9 level text property elements allowed, levels 0-8.

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ParagraphProperties {
    /// bullet used in the paragraph
    pub bullet: Bullet,

    /// Line Spacing
    ///
    /// This element specifies the vertical line spacing that is to be used within a paragraph. This can be specified in two different ways, percentage spacing and font point spacing.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linespacing?view=openxml-3.0.1
    pub line_spacing: SpacingTypeValues,

    /// Spacing after the paragraph
    ///
    /// This element specifies the amount of vertical white space that is present after a paragraph.
    /// This can be specified in two different ways, percentage spacing and font point spacing.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spaceafter?view=openxml-3.0.1
    pub space_after: SpacingTypeValues,

    /// Spacing before the paragraph
    ///
    /// This element specifies the amount of vertical white space that is present before a paragraph.
    /// This can be specified in two different ways, percentage spacing and font point spacing.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacebefore?view=openxml-3.0.1
    pub space_before: SpacingTypeValues,

    /// Text Alignment
    ///
    /// * Center
    /// * Distributed
    /// * Justified
    /// * JustifiedLow
    /// * Left
    /// * Right
    /// * ThaiDistributed
    pub text_alignment: TextAlignmentTypeValues,

    /// FontAlignment
    ///
    /// * Automatic
    /// * Baseline
    /// * Bottom
    /// * Center
    /// * Top
    pub font_alignment: TextFontAlignmentValues,

    /// Indent
    ///
    /// defualt to 0 pt
    pub indent: f64,

    /// Indent Level
    ///
    /// This type specifies the indent level type.
    ///
    /// - a minimum value of greater than or equal to 0.
    /// - a maximum value of less than or equal to 8.
    ///
    /// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextIndentLevelTy_topic_ID0EMGROB.html
    pub indent_level: u64,

    /// LeftMargin
    ///
    /// defualt to 0 pt
    pub left_margin: f64,

    /// RightMargin
    ///
    /// defualt to 0 pt
    pub right_margin: f64,

    /// DefaultTabSize
    ///
    /// default to 72pt (1 inch)
    pub default_tab_size: f64,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tabstoplist?view=openxml-3.0.1
    ///
    /// This element specifies the list of all tab stops that are to be used within a paragraph.
    /// These tabs should be used when describing any custom tab stops within the document.
    /// If these are not specified then the default tab stops of the generating application should be used.
    pub custom_tab_list: Vec<TabStop>,

    /// EastAsianLineBreak
    ///
    /// default to true
    pub east_asian_line_break: bool,

    /// LatinLineBreak
    ///
    /// Allow latin text to wrap in the middle of a word
    ///
    /// default to false
    pub latin_line_break: bool,

    /// Allow hanging punctuation
    ///
    /// default to true
    pub hanging_punctuation: bool,

    /// RightToLeft
    ///
    ///default to false
    pub right_to_left: bool,
}

impl ParagraphProperties {
    pub(crate) fn default() -> Self {
        Self {
            bullet: Bullet::default(),
            line_spacing: SpacingTypeValues::defualt(),
            space_after: SpacingTypeValues::defualt(),
            space_before: SpacingTypeValues::defualt(),
            custom_tab_list: Vec::new(),
            text_alignment: TextAlignmentTypeValues::default(),
            default_tab_size: 72.0,
            east_asian_line_break: true,
            font_alignment: TextFontAlignmentValues::default(),
            hanging_punctuation: true,
            indent: 0.0,
            indent_level: 0,
            latin_line_break: false,
            left_margin: 0.0,
            right_margin: 0.0,
            right_to_left: false,
        }
    }
    pub(crate) fn from_raw(
        raw: Option<XlsxParagraphProperties>,
        default_properties: Option<Self>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        let default = if let Some(default) = default_properties {
            default
        } else {
            Self::default()
        };
        let Some(raw) = raw else { return default };

        let custom_tab_list: Vec<TabStop> = if let Some(tab_list) = raw.clone().tab_list {
            tab_list.into_iter().map(|s| TabStop::from_raw(s)).collect()
        } else {
            default.custom_tab_list
        };

        return Self {
            bullet: Bullet::from_paragraph_properties(
                raw.clone(),
                Some(default.bullet),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
            ),
            line_spacing: if let Some(val) = raw.clone().line_spacing {
                SpacingTypeValues::from_raw(Some(val))
            } else {
                default.line_spacing
            },
            space_after: if let Some(val) = raw.clone().space_after {
                SpacingTypeValues::from_raw(Some(val))
            } else {
                default.space_after
            },
            space_before: if let Some(val) = raw.clone().space_before {
                SpacingTypeValues::from_raw(Some(val))
            } else {
                default.space_before
            },
            text_alignment: if let Some(val) = raw.clone().text_alignment {
                TextAlignmentTypeValues::from_string(Some(val))
            } else {
                default.text_alignment
            },
            font_alignment: if let Some(val) = raw.clone().text_alignment {
                TextFontAlignmentValues::from_string(Some(val))
            } else {
                default.font_alignment
            },
            indent: if let Some(val) = raw.clone().indent {
                emu_to_pt(val)
            } else {
                default.indent
            },
            indent_level: raw.clone().indent_level.unwrap_or(default.indent_level),
            left_margin: if let Some(val) = raw.clone().left_margin {
                emu_to_pt(val)
            } else {
                default.left_margin
            },
            right_margin: if let Some(val) = raw.clone().right_margin {
                emu_to_pt(val)
            } else {
                default.right_margin
            },
            default_tab_size: if let Some(val) = raw.clone().default_tab_size {
                emu_to_pt(val)
            } else {
                default.default_tab_size
            },
            custom_tab_list,
            east_asian_line_break: raw
                .clone()
                .east_asian_line_break
                .unwrap_or(default.east_asian_line_break),
            latin_line_break: raw
                .clone()
                .latin_line_break
                .unwrap_or(default.latin_line_break),
            hanging_punctuation: raw
                .clone()
                .hanging_punctuation
                .unwrap_or(default.hanging_punctuation),
            right_to_left: raw.clone().right_to_left.unwrap_or(default.right_to_left),
        };
    }
}
