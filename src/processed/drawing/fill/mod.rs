pub mod blip_fill;
pub mod fill_rectangle;
pub mod gradient_fill;
pub mod pattern_fill;
pub mod preset_pattern_values;
pub mod tile;
pub mod tile_flip;

#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    raw::drawing::{
        fill::XlsxFillStyleEnum,
        line::outline::XlsxOutline,
        scheme::color_scheme::XlsxColorScheme,
        shape::{
            shape_properties::XlsxShapeProperties,
            visual_group_shape_properties::XlsxVisualGroupShapeProperties,
        },
        text::default_text_run_properties::XlsxTextRunProperties,
    },
};

use super::text::text_run_properties::TextRunProperties;
use blip_fill::BlipFill;
use gradient_fill::GradientFill;
use pattern_fill::PatternFill;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fill?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum Fill {
    // SolidFill(XlsxSolidFill),
    SolidFill(HexColor),

    /// GradientFill
    GradientFill(GradientFill),

    /// grpFill (Group Fill)
    ///
    /// This element specifies a group fill.
    /// When specified, this setting indicates that the parent element is part of a group and should inherit the fill properties of the group.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.groupfill?view=openxml-3.0.1
    // GroupFill,

    /// noFill (No Fill)
    NoFill,

    /// pattFill (Pattern Fill)
    ///
    /// Specifies a pattern fill.
    /// A repeated pattern is used to fill the object.
    PatternFill(PatternFill),

    /// blipFill (Picture Fill)
    ///
    /// specifies the type of picture fill that a picture object has.
    ///
    /// Example:
    /// ```
    /// <p:blipFill>
    ///     <a:blip r:embed="rId2"/>
    ///     <a:stretch>
    ///         <a:fillRect b="10000" r="25000"/>
    ///     </a:stretch>
    /// </p:blipFill>
    /// ```
    BlipFill(BlipFill),
}

impl Fill {
    pub(crate) fn from_raw(
        raw: Option<XlsxFillStyleEnum>,
        parent_group_fill: Option<Self>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };

        Some(match raw {
            XlsxFillStyleEnum::SolidFill(fill) => {
                if let Some(hex) = fill.to_hex(color_scheme.clone(), ref_color.clone()) {
                    Self::SolidFill(hex)
                } else {
                    return None;
                }
            }
            XlsxFillStyleEnum::GradientFill(fill) => {
                if let Some(gradient) =
                    GradientFill::from_raw(Some(fill), color_scheme.clone(), ref_color.clone())
                {
                    Self::GradientFill(gradient)
                } else {
                    Self::NoFill
                }
            }
            XlsxFillStyleEnum::GroupFill(_) => parent_group_fill.unwrap_or(Self::NoFill),
            XlsxFillStyleEnum::NoFill(_) => Self::NoFill,
            XlsxFillStyleEnum::PatternFill(fill) => {
                if let Some(pattern) =
                    PatternFill::from_raw(Some(fill), color_scheme.clone(), ref_color.clone())
                {
                    Self::PatternFill(pattern)
                } else {
                    Self::NoFill
                }
            }
            XlsxFillStyleEnum::BlipFill(fill) => {
                if let Some(blip) = BlipFill::from_raw(
                    Some(fill),
                    drawing_relationship,
                    image_bytes,
                    color_scheme.clone(),
                    ref_color.clone(),
                ) {
                    Self::BlipFill(blip)
                } else {
                    Self::NoFill
                }
            }
        })
    }

    pub(crate) fn from_shape_properties(
        raw: XlsxShapeProperties,
        parent_group_fill: Option<Self>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Option<Self> {
        if let Some(fill) = raw.solid_fill {
            if let Some(hex) = fill.to_hex(color_scheme.clone(), None) {
                return Some(Self::SolidFill(hex));
            } else {
                return None;
            }
        }

        if let Some(fill) = raw.gradient_fill {
            if let Some(gradient) = GradientFill::from_raw(Some(fill), color_scheme.clone(), None) {
                return Some(Self::GradientFill(gradient));
            }
        }

        if let Some(_) = raw.group_fill {
            return parent_group_fill;
        }

        if let Some(_) = raw.no_fill {
            return Some(Self::NoFill);
        }

        if let Some(fill) = raw.pattern_fill {
            if let Some(pattern) = PatternFill::from_raw(Some(fill), color_scheme.clone(), None) {
                return Some(Self::PatternFill(pattern));
            }
        }

        if let Some(fill) = raw.blip_fill {
            if let Some(blip) = BlipFill::from_raw(
                Some(fill),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                None,
            ) {
                return Some(Self::BlipFill(blip));
            }
        }

        return None;
    }

    // group shape does not inherit any parent group fill even when nested in other group shape.
    pub(crate) fn from_group_shape_properties(
        raw: XlsxVisualGroupShapeProperties,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Option<Self> {
        if let Some(fill) = raw.solid_fill {
            if let Some(hex) = fill.to_hex(color_scheme.clone(), None) {
                return Some(Self::SolidFill(hex));
            } else {
                return None;
            }
        }

        if let Some(fill) = raw.gradient_fill {
            if let Some(gradient) = GradientFill::from_raw(Some(fill), color_scheme.clone(), None) {
                return Some(Self::GradientFill(gradient));
            }
        }

        if let Some(_) = raw.no_fill {
            return Some(Self::NoFill);
        }

        if let Some(fill) = raw.pattern_fill {
            if let Some(pattern) = PatternFill::from_raw(Some(fill), color_scheme.clone(), None) {
                return Some(Self::PatternFill(pattern));
            }
        }

        if let Some(fill) = raw.blip_fill {
            if let Some(blip) = BlipFill::from_raw(
                Some(fill),
                drawing_relationship,
                image_bytes,
                color_scheme.clone(),
                None,
            ) {
                return Some(Self::BlipFill(blip));
            }
        }

        return None;
    }

    // parent group fill only applies to shape fill but not outline fill
    pub(crate) fn from_outline(
        raw: XlsxOutline,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Self {
        // default theme accent1 color(#156082) with 15% shade
        let default = Self::SolidFill("020b0fff".to_string());

        if let Some(fill) = raw.solid_fill {
            if let Some(hex) = fill.to_hex(color_scheme.clone(), ref_color.clone()) {
                return Self::SolidFill(hex);
            } else {
                return default;
            }
        }

        if let Some(fill) = raw.gradient_fill {
            if let Some(gradient) =
                GradientFill::from_raw(Some(fill), color_scheme.clone(), ref_color.clone())
            {
                return Self::GradientFill(gradient);
            }
        }

        if let Some(_) = raw.no_fill {
            return Self::NoFill;
        }

        if let Some(fill) = raw.pattern_fill {
            if let Some(pattern) = PatternFill::from_raw(Some(fill), color_scheme.clone(), None) {
                return Self::PatternFill(pattern);
            }
        }

        return default;
    }

    // parent group fill only applies to shape fill but not to text fill (text color)
    pub(crate) fn from_text_run_properties(
        raw: XlsxTextRunProperties,
        default_properties: Option<TextRunProperties>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Self {
        // default: txt1
        let default = Self::SolidFill(ref_color.clone().unwrap_or("000000ff".to_string()));

        if let Some(fill) = raw.blip_fill {
            if let Some(blip) = BlipFill::from_raw(
                Some(fill),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_color.clone(),
            ) {
                return Self::BlipFill(blip);
            }
        }

        if let Some(fill) = raw.gradient_fill {
            if let Some(gradient) =
                GradientFill::from_raw(Some(fill), color_scheme.clone(), ref_color.clone())
            {
                return Self::GradientFill(gradient);
            }
        }

        if let Some(_) = raw.no_fill {
            return Self::NoFill;
        }

        if let Some(fill) = raw.pattern_fill {
            if let Some(pattern) = PatternFill::from_raw(Some(fill), color_scheme.clone(), None) {
                return Self::PatternFill(pattern);
            }
        }

        if let Some(fill) = raw.solid_fill {
            if let Some(hex) = fill.to_hex(color_scheme.clone(), ref_color.clone()) {
                return Self::SolidFill(hex);
            } else {
                return default;
            }
        }

        if let Some(default_properties) = default_properties {
            return default_properties.fill;
        }

        return default;
    }
}
