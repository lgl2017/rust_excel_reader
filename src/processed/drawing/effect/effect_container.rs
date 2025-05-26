#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::raw::drawing::scheme::color_scheme::XlsxColorScheme;
use crate::raw::drawing::st_types::{emu_to_pt, st_percentage_to_float};

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    processed::drawing::fill::Fill,
    raw::drawing::{
        effect::{effect_container::XlsxEffectContainer, effect_list::XlsxEffectList},
        image::blip::XlsxBlip,
    },
};

use super::reflection::Reflection;
use super::relative_offset::RelativeOffset;
use super::soft_edge::SoftEdge;
use super::transform::TransformEffect;
use super::{
    blend::Blend, blur::Blur, color_change::ColorChange,
    effect_reference::EffectReferenceTypeValues, glow::Glow, hue_saturation_luminance::HslEffect,
    inner_shadow::InnerShadow, luminance::Luminance, outer_shadow::OuterShadow,
    preset_shadow::PresetShadow, tint::Tint,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainer?view=openxml-3.0.1
///
/// A list of effects.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct EffectContainer {
    /// alphaBiLevel (Alpha Bi-Level Effect):
    ///
    /// Alpha (Opacity) values less than the threshold are changed to 0 (fully transparent) and alpha values greater than or equal to the threshold are changed to 100% (fully opaque).
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphabilevel?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub alpha_bi_level: Option<f64>,

    /// alphaCeiling (Alpha Ceiling Effect)
    ///
    /// When present, Alpha (opacity) values greater than zero are changed to 100%.
    /// In other words, anything partially opaque becomes fully opaque.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphaceiling?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub apply_alpha_ceiling: Option<bool>,

    /// alphaFloor (Alpha Floor Effect)
    ///
    /// when present, Alpha (opacity) values less than 100% are changed to zero.
    /// In other words, anything partially transparent becomes fully transparent.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphafloor?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub apply_alpha_floor: Option<bool>,

    /// alphaInv (Alpha Inverse Effect)
    ///
    /// This element represents an alpha inverse effect.
    /// Alpha (opacity) values are inverted by subtracting from 100%.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphainverse?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub alpha_inverse: Option<HexColor>,

    /// alphaMod (Alpha Modulate Effect)
    ///
    /// This element represents an alpha modulate effect.
    /// Effect alpha (opacity) values are multiplied by a fixed percentage.
    /// The effect container specifies an effect containing alpha values to modulate
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphamodulationeffect?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub alpha_modulation: Option<Box<EffectContainer>>,

    /// alphaModFix (Alpha Modulate Fixed Effect)
    ///
    /// This element represents an alpha modulate fixed effect.
    /// Effect alpha (opacity) values are multiplied by a fixed percentage
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphamodulationfixed?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub alpha_modulation_fixed: Option<f64>,

    // alphaOutset (Alpha Inset/Outset Effect)
    ///
    /// This is equivalent to an alpha ceiling, followed by alpha blur, followed by either an alpha ceiling (positive radius) or alpha floor (negative radius).
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphaoutset?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub alpha_outset: Option<f64>,

    /// alphaRepl (Alpha Replace Effect)
    ///
    /// This element specifies an alpha replace effect.
    /// Effect alpha (opacity) values are replaced by a fixed alpha.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphareplace?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub alpha_replace: Option<f64>,

    /// biLevel (Bi-Level (Black/White) Effect)
    ///
    /// This element specifies a bi-level (black/white) effect.
    /// Input colors whose luminance is less than the specified threshold value are changed to black.
    /// Input colors whose luminance are greater than or equal the specified value are set to white.
    /// The alpha effect values are unaffected by this effect.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphabilevel?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub bi_level: Option<f64>,

    /// blend (Blend Effect)
    ///
    /// Specifies a blend of several effects.
    /// The container specifies the raw effects to blend while the blend mode specifies how the effects are to be blended.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blend?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub blend: Option<Blend>,

    /// blur (Blur Effect)
    ///
    /// a blur effect that is applied to the entire shape, including its fill.
    /// All color channels, including alpha, are affected.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blur?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub blur: Option<Blur>,

    /// clrChange (Color Change Effect)
    ///
    /// This element specifies a Color Change Effect.
    /// Instances of clrFrom are replaced with instances of clrTo
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorchange?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub color_change: Option<ColorChange>,

    /// clrRepl (Solid Color Replacement)
    ///
    /// specifies a solid color replacement value.
    /// All effect colors are changed to a fixed color. Alpha values are unaffected.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorreplacement?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub solid_color_replacement: Option<HexColor>,

    /// cont (Effect Container)
    ///
    /// A list of effects.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainer?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub effect_container: Option<Box<EffectContainer>>,

    /// duotone (Duotone Effect)
    ///
    /// This element specifies a duotone effect.
    /// For each pixel, combines clr1 and clr2 through a linear interpolation to determine the new color for that pixel.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.duotone?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub duotone: Option<HexColor>,

    /// effect (Effect)
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effect?view=openxml-3.0.1
    ///
    /// This element specifies a reference to an existing effect container
    ///
    /// * Container: refer to an effect container with the name specified
    /// * Fill: refers to the fill effect
    /// * Line: refers to the line effect
    /// * FillLine: refers to the combined fill and line effects
    /// * Children: refers to the combined effects from logical child shapes or text
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub effect_reference: Option<EffectReferenceTypeValues>,

    /// fill (Fill)
    ///
    /// This element specifies a fill which is one of blipFill, gradFill, grpFill, noFill, pattFill or solidFill.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fill?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub fill: Option<Fill>,

    /// fillOverlay (Fill Overlay Effect)
    ///
    ///  specifies a fill overlay effect.
    /// A fill overlay can be used to specify an additional fill for an object and blend the two fills together
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.filloverlay?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub fill_overlay: Option<Box<Fill>>,

    /// glow (Glow Effect)
    ///
    /// specifies a glow effect, in which a color blurred outline is added outside the edges of the object.
    ///
    /// Example:
    /// ```
    /// <a:glow rad="10">
    ///     <a:schemeClr val="phClr">
    ///         <a:lumMod val="99000" />
    ///         <a:satMod val="120000" />
    ///          a:shade val="78000" />
    ///     </a:schemeClr>
    /// </a:glow>
    /// ```
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.glow?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub glow: Option<Glow>,

    /// grayscl (Gray Scale Effect)
    ///
    /// When present, Converts all effect color values to a shade of gray, corresponding to their luminance.
    /// Effect alpha (opacity) values are unaffected.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.grayscale?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub gray_scale: Option<bool>,

    /// hsl (Hue Saturation Luminance Effect)
    ///
    /// This element specifies a hue/saturation/luminance effect.
    /// The hue, saturation, and luminance can each be adjusted relative to its current value.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hsl?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub hsl: Option<HslEffect>,

    /// innerShdw (Inner Shadow Effect)
    ///
    /// specifies an inner shadow effect.
    /// A shadow is applied within the edges of the object according to the parameters given by the attributes
    ///
    ///  Example:
    /// ```
    /// <a:innerShdw blurRad="10" dir"90" dist="10">
    ///     <a:schemeClr val="phClr">
    ///         <a:lumMod val="99000" />
    ///         <a:satMod val="120000" />
    ///          a:shade val="78000" />
    ///     </a:schemeClr>
    /// </a:innerShdw>
    /// ```
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.innershadow?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub innder_shadow: Option<InnerShadow>,

    /// lum (Luminance Effect)
    ///
    /// This element specifies a luminance effect.
    /// Brightness linearly shifts all colors closer to white or black.
    /// Contrast scales all colors to be either closer or further apart.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminanceeffect?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub luminance: Option<Luminance>,

    /// outerShdw (Outer Shadow Effect)
    ///
    /// specifies an outer shadow effect.
    ///
    ///  Example:
    /// ```
    /// <a:outerShdw blurRad="57150" dist="38100" dir="5400000" algn="ctr" rotWithShape="0" >
    ///     <a:schemeClr val="phClr">
    ///         <a:lumMod val="99000" />
    ///         <a:satMod val="120000" />
    ///          a:shade val="78000" />
    ///     </a:schemeClr>
    /// </a:outerShdw>
    /// ```
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outershadow?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub outer_shadow: Option<OuterShadow>,

    /// prstShdw (Preset Shadow)
    ///
    /// specifies that a preset shadow is to be used.
    ///
    /// Each preset shadow is equivalent to a specific outer shadow effect.
    /// For each preset shadow, the color element, direction attribute, and distance attribute represent the color, direction, and distance parameters of the corresponding outer shadow.
    /// Additionally, the rotateWithShape attribute of corresponding outer shadow is always false. Other non-default parameters of the outer shadow are dependent on the prst attribute
    ///
    ///  Example:
    /// ```
    /// <a:prstShdw dir"90" dist="10" prst="shdw19">
    ///     <a:schemeClr val="phClr">
    ///         <a:lumMod val="99000" />
    ///         <a:satMod val="120000" />
    ///          a:shade val="78000" />
    ///     </a:schemeClr>
    /// </a:prstShdw>
    /// ```
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetshadow?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub preset_shadow: Option<PresetShadow>,

    /// reflection (Reflection Effect)
    ///
    /// This element specifies a reflection effect.
    ///
    /// Example:
    /// ```
    /// <a:reflection blurRad="151308" stA="88815" endPos="65000" dist="402621"dir="5400000" sy="-100000" algn="bl" rotWithShape="0" />
    /// ```
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.reflection?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub reflection: Option<Reflection>,

    /// relOff (Relative Offset Effect)
    ///
    /// This element specifies a relative offset effect. Sets up a new origin by offsetting relative to the size of the previous effect.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.relativeoffset?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub relative_offset: Option<RelativeOffset>,

    /// softEdge (Soft Edge Effect)
    ///
    /// This element specifies a soft edge effect.
    /// The edges of the shape are blurred, while the fill is not affected.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.softedge?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub soft_edge: Option<SoftEdge>,

    /// tint (Tint Effect)
    ///
    /// This element specifies a tint effect.
    /// Shifts effect color values towards/away from hue by the specified amount.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tinteffect?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub tint: Option<Tint>,

    /// xfrm (Transform Effect)
    ///
    /// This element specifies a transform effect. The transform is applied to each point in the shape's geometry using the following matrix:
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transformeffect?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub transform: Option<TransformEffect>,

    /// name (Name)
    ///
    /// Specifies an optional name for this list of effects, so that it can be referred to later.
    /// Shall be unique across all effect trees and effect containers.
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub name: Option<String>,

    /// type (Effect Container Type)
    ///
    /// Specifies the kind of container, either sibling or tree.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainervalues?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub r#type: Option<EffectContainerTypeValues>,
}

impl EffectContainer {
    pub(crate) fn from_raw(
        raw: Option<Box<XlsxEffectContainer>>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Box<Self>> {
        let Some(raw) = raw else { return None };

        return Some(Box::new(Self {
            alpha_bi_level: if let Some(v) = raw.clone().alpha_bi_level {
                Self::percentage_int_to_float_helper(v.thresh)
            } else {
                None
            },
            apply_alpha_ceiling: raw.clone().alpha_ceiling,
            apply_alpha_floor: raw.clone().alpha_floor,
            alpha_inverse: if let Some(v) = raw.clone().alpha_inv {
                v.to_hex(color_scheme.clone(), ref_color.clone())
            } else {
                None
            },
            alpha_modulation: if let Some(v) = raw.clone().alpha_mod {
                EffectContainer::from_raw(
                    v.cont,
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_color.clone(),
                )
            } else {
                None
            },
            alpha_modulation_fixed: if let Some(v) = raw.clone().alpha_mod_fix {
                Self::percentage_int_to_float_helper(v.amt)
            } else {
                None
            },
            alpha_outset: if let Some(v) = raw.clone().alpha_outset {
                Self::emu_to_pt_helper(v.rad)
            } else {
                None
            },
            alpha_replace: if let Some(v) = raw.clone().alpha_repl {
                Self::percentage_unsigned_int_to_float_helper(v.a)
            } else {
                None
            },
            bi_level: if let Some(v) = raw.clone().bi_level {
                Self::percentage_int_to_float_helper(v.thresh)
            } else {
                None
            },
            blend: Blend::from_raw(
                raw.clone().blend,
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_color.clone(),
            ),
            blur: Blur::from_raw(raw.clone().blur),
            color_change: ColorChange::from_raw(
                raw.clone().clr_change,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            solid_color_replacement: if let Some(v) = raw.clone().clr_repl {
                v.to_hex(color_scheme.clone(), ref_color.clone())
            } else {
                None
            },
            effect_container: EffectContainer::from_raw(
                raw.clone().cont,
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_color.clone(),
            ),
            effect_reference: EffectReferenceTypeValues::from_raw(raw.clone().effect),
            duotone: if let Some(v) = raw.clone().duotone {
                v.to_hex(color_scheme.clone(), ref_color.clone())
            } else {
                None
            },
            fill: if let Some(v) = raw.clone().fill {
                Fill::from_raw(
                    Some(v),
                    None,
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_color.clone(),
                )
            } else {
                None
            },
            fill_overlay: if let Some(v) = raw.clone().fill_overlay {
                if let Some(fill) = Fill::from_raw(
                    Some(*v),
                    None,
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_color.clone(),
                ) {
                    Some(Box::new(fill))
                } else {
                    None
                }
            } else {
                None
            },
            gray_scale: raw.clone().grayscl,
            glow: Glow::from_raw(raw.clone().glow, color_scheme.clone(), ref_color.clone()),
            hsl: HslEffect::from_raw(raw.clone().hsl),
            innder_shadow: InnerShadow::from_raw(
                raw.clone().innder_shadow,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            luminance: Luminance::from_raw(raw.clone().lum),
            outer_shadow: OuterShadow::from_raw(
                raw.clone().outer_shadow,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            preset_shadow: PresetShadow::from_raw(
                raw.clone().preset_shadow,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            reflection: Reflection::from_raw(raw.clone().reflection),
            relative_offset: RelativeOffset::from_raw(raw.clone().relative_offset),
            soft_edge: SoftEdge::from_raw(raw.clone().soft_edge),
            tint: Tint::from_raw(raw.clone().tint),
            transform: TransformEffect::from_raw(raw.clone().transform),
            name: raw.clone().name,
            r#type: Some(EffectContainerTypeValues::from_raw(raw.clone().r#type)),
        }));
    }

    pub(crate) fn from_raw_blip(
        raw: XlsxBlip,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Box<Self> {
        return Box::new(Self {
            alpha_bi_level: if let Some(v) = raw.clone().alpha_bi_level {
                Self::percentage_int_to_float_helper(v.thresh)
            } else {
                None
            },
            apply_alpha_ceiling: raw.clone().alpha_ceiling,
            apply_alpha_floor: raw.clone().alpha_floor,
            alpha_inverse: if let Some(v) = raw.clone().alpha_inv {
                v.to_hex(color_scheme.clone(), ref_color.clone())
            } else {
                None
            },
            alpha_modulation: if let Some(v) = raw.clone().alpha_mod {
                EffectContainer::from_raw(
                    v.cont,
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_color.clone(),
                )
            } else {
                None
            },
            alpha_modulation_fixed: if let Some(v) = raw.clone().alpha_mod_fix {
                Self::percentage_int_to_float_helper(v.amt)
            } else {
                None
            },
            alpha_outset: None,
            alpha_replace: if let Some(v) = raw.clone().alpha_repl {
                Self::percentage_unsigned_int_to_float_helper(v.a)
            } else {
                None
            },
            bi_level: if let Some(v) = raw.clone().bi_level {
                Self::percentage_int_to_float_helper(v.thresh)
            } else {
                None
            },
            blend: None,
            blur: Blur::from_raw(raw.clone().blur),
            color_change: ColorChange::from_raw(
                raw.clone().clr_change,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            solid_color_replacement: if let Some(v) = raw.clone().clr_repl {
                v.to_hex(color_scheme.clone(), ref_color.clone())
            } else {
                None
            },
            effect_container: None,
            effect_reference: None,
            duotone: if let Some(v) = raw.clone().duotone {
                v.to_hex(color_scheme.clone(), ref_color.clone())
            } else {
                None
            },
            fill: None,
            fill_overlay: if let Some(v) = raw.clone().fill_overlay {
                if let Some(fill) = Fill::from_raw(
                    Some(*v),
                    None,
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_color.clone(),
                ) {
                    Some(Box::new(fill))
                } else {
                    None
                }
            } else {
                None
            },
            gray_scale: raw.clone().grayscl,
            glow: None,
            hsl: HslEffect::from_raw(raw.clone().hsl),
            innder_shadow: None,
            luminance: Luminance::from_raw(raw.clone().lum),
            outer_shadow: None,
            preset_shadow: None,
            reflection: None,
            relative_offset: None,
            soft_edge: None,
            tint: Tint::from_raw(raw.clone().tint),
            transform: None,
            name: None,
            r#type: None,
        });
    }

    pub(crate) fn from_raw_effect_list(
        raw: Option<XlsxEffectList>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Box<Self>> {
        let Some(raw) = raw else { return None };

        return Some(Box::new(Self {
            alpha_bi_level: None,
            apply_alpha_ceiling: None,
            apply_alpha_floor: None,
            alpha_inverse: None,
            alpha_modulation: None,
            alpha_modulation_fixed: None,
            alpha_outset: None,
            alpha_replace: None,
            bi_level: None,
            blend: None,
            blur: Blur::from_raw(raw.clone().blur),
            color_change: None,
            solid_color_replacement: None,
            effect_container: None,
            effect_reference: None,
            duotone: None,
            fill: None,
            fill_overlay: if let Some(v) = raw.clone().fill_overlay {
                if let Some(fill) = Fill::from_raw(
                    Some(*v),
                    None,
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_color.clone(),
                ) {
                    Some(Box::new(fill))
                } else {
                    None
                }
            } else {
                None
            },
            gray_scale: None,
            glow: Glow::from_raw(raw.clone().glow, color_scheme.clone(), ref_color.clone()),
            hsl: None,
            innder_shadow: InnerShadow::from_raw(
                raw.clone().innder_shadow,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            luminance: None,
            outer_shadow: OuterShadow::from_raw(
                raw.clone().outer_shadow,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            preset_shadow: PresetShadow::from_raw(
                raw.clone().preset_shadow,
                color_scheme.clone(),
                ref_color.clone(),
            ),
            reflection: Reflection::from_raw(raw.clone().reflection),
            relative_offset: None,
            soft_edge: SoftEdge::from_raw(raw.clone().soft_edge),
            tint: None,
            transform: None,
            name: None,
            r#type: None,
        }));
    }

    fn percentage_int_to_float_helper(i: Option<i64>) -> Option<f64> {
        let Some(i) = i else { return None };
        return Some(st_percentage_to_float(i));
    }

    fn percentage_unsigned_int_to_float_helper(u: Option<u64>) -> Option<f64> {
        let Some(u) = u else { return None };
        return Some(st_percentage_to_float(u as i64));
    }

    fn emu_to_pt_helper(i: Option<i64>) -> Option<f64> {
        let Some(i) = i else { return None };
        return Some(emu_to_pt(i));
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainervalues?view=openxml-3.0.1
///
/// * Sibling
/// * Tree
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum EffectContainerTypeValues {
    Sibling,
    Tree,
}

impl EffectContainerTypeValues {
    pub(crate) fn from_raw(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::Tree };
        return match s.as_ref() {
            "sib" => Self::Sibling,
            "tree" => Self::Tree,
            _ => Self::Tree,
        };
    }
}
