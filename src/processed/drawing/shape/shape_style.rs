use std::collections::BTreeMap;

#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    packaging::relationship::XlsxRelationships,
    processed::drawing::{effect::effect_style::EffectStyle, fill::Fill, line::outline::Outline},
    raw::drawing::{
        scheme::color_scheme::XlsxColorScheme,
        shape::{
            shape_properties::XlsxShapeProperties,
            visual_group_shape_properties::XlsxVisualGroupShapeProperties,
        },
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapestyle?view=openxml-3.0.1
///
/// This element specifies the style information for a shape.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ShapeStyle {
    /// fill
    pub fill: Fill,

    /// ln (Outline)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub outline: Option<Outline>,

    /// effects:
    /// - 2D effects (glow, shadow, etc.)
    /// - 3D Scene effects
    /// - 3D shape effects
    pub effects: EffectStyle,
}

impl ShapeStyle {
    pub(crate) fn default() -> Self {
        Self {
            fill: Fill::NoFill,
            outline: None,
            effects: EffectStyle {
                scene3d_properties: None,
                shape3d_properties: None,
                effect_list: None,
            },
        }
    }

    pub(crate) fn from_shape_properties(
        raw: XlsxShapeProperties,
        parent_group_fill: Option<Fill>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        line_ref: Option<Outline>,
        fill_ref: Option<Fill>,
        effect_ref: Option<EffectStyle>,
    ) -> Self {
        let fill = Fill::from_shape_properties(
            raw.clone(),
            parent_group_fill.clone(),
            drawing_relationship.clone(),
            image_bytes.clone(),
            color_scheme.clone(),
        );

        let fill = if let Some(fill) = fill {
            Some(fill)
        } else {
            fill_ref
        };

        let outline = Outline::with_reference(raw.clone().outline, line_ref, color_scheme.clone());

        return Self {
            fill: fill.unwrap_or(Fill::NoFill),
            outline,
            effects: EffectStyle::from_shape_properties(
                raw.clone(),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                effect_ref.clone(),
            ),
        };
    }

    // group shape does not inherit any parent group fill even when nested in other group shape.
    pub(crate) fn from_group_shape_properties(
        raw: XlsxVisualGroupShapeProperties,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        return Self {
            // group shape itself does not have any fill.
            // Fill specified in the property is to be used for its children contents such as shapes
            fill: Fill::NoFill,
            outline: None,
            effects: EffectStyle::from_group_shape_properties(
                raw.clone(),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
            ),
        };
    }
}
