#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    processed::drawing::{
        scene::scene_3d_properties::Scene3DProperties,
        shape::shape_3d_properties::Shape3DProperties,
    },
    raw::drawing::{
        effect::effect_style::XlsxEffectStyle,
        scheme::color_scheme::XlsxColorScheme,
        shape::{
            shape_properties::XlsxShapeProperties,
            visual_group_shape_properties::XlsxVisualGroupShapeProperties,
        },
    },
};

use super::effect_container::EffectContainer;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectstyle?view=openxml-3.0.1
///
/// This element defines a set of effects and 3D properties that can be applied to an object.
///
/// Example:
/// ```
/// <effectStyle>
///   <effectLst>
///     <outerShdw blurRad="57150" dist="38100" dir="5400000" algn="ctr" rotWithShape="0">
///       <schemeClr val="phClr">
///         <shade val="9000"/>
///         <satMod val="105000"/>
///         <alpha val="48000"/>
///       </schemeClr>
///     </outerShdw>
///   </effectLst>
/// </effectStyle>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct EffectStyle {
    /// scene3d (3D Scene Properties)
    // #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub scene3d_properties: Option<Scene3DProperties>,

    /// sp3d (Apply 3D shape properties)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub shape3d_properties: Option<Shape3DProperties>,

    /// a list of effect to apply to the shape
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub effect_list: Option<Box<EffectContainer>>,
}

impl EffectStyle {
    pub(crate) fn from_raw(
        raw: Option<XlsxEffectStyle>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else {
            return None;
        };
        return Some(Self {
            scene3d_properties: Scene3DProperties::from_raw(raw.clone().scene3d),
            shape3d_properties: Shape3DProperties::from_raw(
                raw.clone().shape3d,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            effect_list: EffectContainer::from_raw_effect_list(
                raw.clone().effect_lst,
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_color.clone(),
            ),
        });
    }

    pub(crate) fn from_shape_properties(
        raw: XlsxShapeProperties,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        effect_ref: Option<EffectStyle>,
    ) -> Self {
        let scene = if let Some(scene) = raw.clone().scene3d {
            Scene3DProperties::from_raw(Some(scene))
        } else {
            if let Some(effect_ref) = effect_ref.clone() {
                effect_ref.scene3d_properties
            } else {
                None
            }
        };

        let shape = if let Some(shape) = raw.clone().shape3d {
            Shape3DProperties::from_raw(Some(shape), color_scheme.clone(), None)
        } else {
            if let Some(effect_ref) = effect_ref.clone() {
                effect_ref.shape3d_properties
            } else {
                None
            }
        };

        let effect = if let Some(effect) = raw.clone().effect_list {
            EffectContainer::from_raw_effect_list(
                Some(effect),
                drawing_relationship,
                image_bytes.clone(),
                color_scheme.clone(),
                None,
            )
        } else {
            if let Some(effect_ref) = effect_ref.clone() {
                effect_ref.effect_list
            } else {
                None
            }
        };

        return Self {
            scene3d_properties: scene,
            shape3d_properties: shape,
            effect_list: effect,
        };
    }

    pub(crate) fn from_group_shape_properties(
        raw: XlsxVisualGroupShapeProperties,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        let scene = if let Some(scene) = raw.clone().scene3d {
            Scene3DProperties::from_raw(Some(scene))
        } else {
            None
        };

        let effect = if let Some(effect) = raw.clone().effect_list {
            EffectContainer::from_raw_effect_list(
                Some(effect),
                drawing_relationship,
                image_bytes.clone(),
                color_scheme.clone(),
                None,
            )
        } else {
            None
        };

        return Self {
            scene3d_properties: scene,
            shape3d_properties: None,
            effect_list: effect,
        };
    }
}
