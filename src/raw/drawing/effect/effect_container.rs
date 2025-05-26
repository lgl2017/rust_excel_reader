use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::excel::XmlReader;

use super::{
    alpha_bi_level::XlsxAlphaBiLevel, alpha_ceiling::XlsxAlphaCeiling, alpha_floor::XlsxAlphaFloor,
    alpha_inverse::XlsxAlphaInverse, alpha_modulation::XlsxAlphaModulation,
    alpha_modulation_fixed::XlsxAlphaModulationFixed, alpha_outset::XlsxAlphaOutset,
    alpha_replace::XlsxAlphaReplace, bi_level::XlsxBiLevel, blend::XlsxBlend, blur::XlsxBlur,
    color_change::XlsxColorChange, color_replacement::XlsxColorReplacement, duotone::XlsxDuotone,
    effect::XlsxEffect, fill::XlsxFill, fill_overlay::XlsxFillOverlay, glow::XlsxGlow,
    gray_scale::XlsxGrayScale, hue_saturation_luminance::XlsxHsl, inner_shadow::XlsxInnerShadow,
    luminance::XlsxLuminance, outer_shadow::XlsxOuterShadow, preset_shadow::XlsxPresetShadow,
    reflection::XlsxReflection, relative_offset::XlsxRelativeOffset, soft_edge::XlsxSoftEdge,
    tint::XlsxTint, transform::XlsxTransformEffect,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectdag?view=openxml-3.0.1
pub type XlsxEffectDag = XlsxEffectContainer;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainer?view=openxml-3.0.1
///
/// A list of effects.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxEffectContainer {
    // Child Elements	Subclause
    /// alphaBiLevel (Alpha Bi-Level Effect):
    ///
    /// Alpha (Opacity) values less than the threshold are changed to 0 (fully transparent) and alpha values greater than or equal to the threshold are changed to 100% (fully opaque).
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphabilevel?view=openxml-3.0.1
    pub alpha_bi_level: Option<XlsxAlphaBiLevel>,

    /// alphaCeiling (Alpha Ceiling Effect)
    ///
    /// When present, Alpha (opacity) values greater than zero are changed to 100%.
    /// In other words, anything partially opaque becomes fully opaque.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphaceiling?view=openxml-3.0.1
    pub alpha_ceiling: Option<XlsxAlphaCeiling>,

    /// alphaFloor (Alpha Floor Effect)
    ///
    /// when present, Alpha (opacity) values less than 100% are changed to zero.
    /// In other words, anything partially transparent becomes fully transparent.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphafloor?view=openxml-3.0.1
    pub alpha_floor: Option<XlsxAlphaFloor>,

    /// alphaInv (Alpha Inverse Effect)
    ///
    /// This element represents an alpha inverse effect.
    /// Alpha (opacity) values are inverted by subtracting from 100%.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphainverse?view=openxml-3.0.1
    pub alpha_inv: Option<XlsxAlphaInverse>,

    /// alphaMod (Alpha Modulate Effect)
    ///
    /// This element represents an alpha modulate effect.
    /// Effect alpha (opacity) values are multiplied by a fixed percentage.
    /// The effect container specifies an effect containing alpha values to modulate
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphamodulationeffect?view=openxml-3.0.1
    pub alpha_mod: Option<XlsxAlphaModulation>,

    /// alphaModFix (Alpha Modulate Fixed Effect)
    ///
    /// This element represents an alpha modulate fixed effect.
    /// Effect alpha (opacity) values are multiplied by a fixed percentage
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphamodulationfixed?view=openxml-3.0.1
    pub alpha_mod_fix: Option<XlsxAlphaModulationFixed>,

    // alphaOutset (Alpha Inset/Outset Effect)
    ///
    /// This is equivalent to an alpha ceiling, followed by alpha blur, followed by either an alpha ceiling (positive radius) or alpha floor (negative radius).
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphaoutset?view=openxml-3.0.1
    pub alpha_outset: Option<XlsxAlphaOutset>,

    /// alphaRepl (Alpha Replace Effect)
    ///
    /// This element specifies an alpha replace effect.
    /// Effect alpha (opacity) values are replaced by a fixed alpha.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphareplace?view=openxml-3.0.1
    pub alpha_repl: Option<XlsxAlphaReplace>,

    /// biLevel (Bi-Level (Black/White) Effect)
    ///
    /// This element specifies a bi-level (black/white) effect.
    /// Input colors whose luminance is less than the specified threshold value are changed to black.
    /// Input colors whose luminance are greater than or equal the specified value are set to white.
    /// The alpha effect values are unaffected by this effect.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphabilevel?view=openxml-3.0.1
    pub bi_level: Option<XlsxBiLevel>,

    /// blend (Blend Effect)
    ///
    /// Specifies a blend of several effects.
    /// The container specifies the raw effects to blend while the blend mode specifies how the effects are to be blended.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blend?view=openxml-3.0.1
    pub blend: Option<XlsxBlend>,

    /// blur (Blur Effect)
    ///
    /// a blur effect that is applied to the entire shape, including its fill.
    /// All color channels, including alpha, are affected.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blur?view=openxml-3.0.1
    pub blur: Option<XlsxBlur>,

    /// clrChange (Color Change Effect)
    ///
    /// This element specifies a Color Change Effect.
    /// Instances of clrFrom are replaced with instances of clrTo
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorchange?view=openxml-3.0.1
    pub clr_change: Option<XlsxColorChange>,

    /// clrRepl (Solid Color Replacement)
    ///
    /// specifies a solid color replacement value.
    /// All effect colors are changed to a fixed color. Alpha values are unaffected.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorreplacement?view=openxml-3.0.1
    pub clr_repl: Option<XlsxColorReplacement>,

    /// cont (Effect Container)
    ///
    /// A list of effects.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainer?view=openxml-3.0.1
    pub cont: Option<Box<XlsxEffectContainer>>,

    /// duotone (Duotone Effect)
    ///
    /// This element specifies a duotone effect.
    /// For each pixel, combines clr1 and clr2 through a linear interpolation to determine the new color for that pixel.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.duotone?view=openxml-3.0.1
    pub duotone: Option<XlsxDuotone>,

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
    pub effect: Option<XlsxEffect>,

    /// fill (Fill)
    ///
    /// This element specifies a fill which is one of blipFill, gradFill, grpFill, noFill, pattFill or solidFill.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fill?view=openxml-3.0.1
    pub fill: Option<XlsxFill>,

    /// fillOverlay (Fill Overlay Effect)
    ///
    ///  specifies a fill overlay effect.
    /// A fill overlay can be used to specify an additional fill for an object and blend the two fills together
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.filloverlay?view=openxml-3.0.1
    pub fill_overlay: Option<Box<XlsxFillOverlay>>,

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
    pub glow: Option<XlsxGlow>,

    /// grayscl (Gray Scale Effect)
    ///
    /// When present, Converts all effect color values to a shade of gray, corresponding to their luminance.
    /// Effect alpha (opacity) values are unaffected.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.grayscale?view=openxml-3.0.1
    pub grayscl: Option<XlsxGrayScale>,

    /// hsl (Hue Saturation Luminance Effect)
    ///
    /// This element specifies a hue/saturation/luminance effect.
    /// The hue, saturation, and luminance can each be adjusted relative to its current value.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hsl?view=openxml-3.0.1
    pub hsl: Option<XlsxHsl>,

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
    pub innder_shadow: Option<XlsxInnerShadow>,

    /// lum (Luminance Effect)
    ///
    /// This element specifies a luminance effect.
    /// Brightness linearly shifts all colors closer to white or black.
    /// Contrast scales all colors to be either closer or further apart.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminanceeffect?view=openxml-3.0.1
    pub lum: Option<XlsxLuminance>,

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
    pub outer_shadow: Option<XlsxOuterShadow>,

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
    pub preset_shadow: Option<XlsxPresetShadow>,

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
    pub reflection: Option<XlsxReflection>,

    /// relOff (Relative Offset Effect)
    ///
    /// This element specifies a relative offset effect. Sets up a new origin by offsetting relative to the size of the previous effect.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.relativeoffset?view=openxml-3.0.1
    pub relative_offset: Option<XlsxRelativeOffset>,

    /// softEdge (Soft Edge Effect)
    ///
    /// This element specifies a soft edge effect.
    /// The edges of the shape are blurred, while the fill is not affected.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.softedge?view=openxml-3.0.1
    pub soft_edge: Option<XlsxSoftEdge>,

    /// tint (Tint Effect)
    ///
    /// This element specifies a tint effect.
    /// Shifts effect color values towards/away from hue by the specified amount.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tinteffect?view=openxml-3.0.1
    pub tint: Option<XlsxTint>,

    /// xfrm (Transform Effect)
    ///
    /// This element specifies a transform effect. The transform is applied to each point in the shape's geometry using the following matrix:
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transformeffect?view=openxml-3.0.1
    pub transform: Option<XlsxTransformEffect>,

    // attributes
    /// name (Name)
    ///
    /// Specifies an optional name for this list of effects, so that it can be referred to later.
    /// Shall be unique across all effect trees and effect containers.
    pub name: Option<String>,

    /// type (Effect Container Type)
    ///
    /// Specifies the kind of container, either sibling or tree.
    /// allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainervalues?view=openxml-3.0.1
    pub r#type: Option<String>,
}

impl XlsxEffectContainer {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        return XlsxEffectContainer::load_helper(reader, e, b"cont");
    }

    pub(crate) fn load_effect_dag(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
    ) -> anyhow::Result<Self> {
        return XlsxEffectContainer::load_helper(reader, e, b"effectDag");
    }

    fn load_helper(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
        tag: &[u8],
    ) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut container = Self {
            alpha_bi_level: None,
            alpha_ceiling: None,
            alpha_floor: None,
            alpha_inv: None,
            alpha_mod: None,
            alpha_mod_fix: None,
            alpha_outset: None,
            alpha_repl: None,
            bi_level: None,
            blend: None,
            blur: None,
            clr_change: None,
            clr_repl: None,
            cont: None,
            duotone: None,
            effect: None,
            fill: None,
            fill_overlay: None,
            glow: None,
            grayscl: None,
            hsl: None,
            innder_shadow: None,
            lum: None,
            outer_shadow: None,
            preset_shadow: None,
            reflection: None,
            relative_offset: None,
            soft_edge: None,
            tint: None,
            transform: None,
            name: None,
            r#type: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"name" => {
                            container.name = Some(string_value);
                        }
                        b"type" => {
                            container.r#type = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaBiLevel" => {
                    container.alpha_bi_level = Some(XlsxAlphaBiLevel::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaCeiling" => {
                    container.alpha_ceiling = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaFloor" => {
                    container.alpha_floor = Some(true);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaInv" => {
                    container.alpha_inv = XlsxAlphaInverse::load(reader, b"alphaInv")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaMod" => {
                    container.alpha_mod = Some(XlsxAlphaModulation::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaModFix" => {
                    container.alpha_mod_fix = Some(XlsxAlphaModulationFixed::load(e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaOutset" => {
                    container.alpha_outset = Some(XlsxAlphaOutset::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaRepl" => {
                    container.alpha_repl = Some(XlsxAlphaReplace::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"biLevel" => {
                    container.bi_level = Some(XlsxBiLevel::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blend" => {
                    container.blend = Some(XlsxBlend::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blur" => {
                    container.blur = Some(XlsxBlur::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrChange" => {
                    container.clr_change = Some(XlsxColorChange::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrRepl" => {
                    container.clr_repl = XlsxColorReplacement::load(reader, b"clrRepl")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cont" => {
                    container.cont = Some(Box::new(XlsxEffectContainer::load(reader, e)?));
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"duotone" => {
                    container.duotone = XlsxDuotone::load(reader, b"duotone")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effect" => {
                    container.effect = Some(XlsxEffect::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                    container.fill = XlsxFill::load(reader, b"fill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillOverlay" => {
                    if let Some(fill_overlay) = XlsxFillOverlay::load(reader, b"fillOverlay")? {
                        container.fill_overlay = Some(Box::new(fill_overlay));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"glow" => {
                    container.glow = Some(XlsxGlow::load(reader, e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grayscl" => {
                    container.grayscl = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hsl" => {
                    container.hsl = Some(XlsxHsl::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"innerShdw" => {
                    container.innder_shadow = Some(XlsxInnerShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lum" => {
                    container.lum = Some(XlsxLuminance::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"outerShdw" => {
                    container.outer_shadow = Some(XlsxOuterShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"prstShdw" => {
                    container.preset_shadow = Some(XlsxPresetShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"reflection" => {
                    container.reflection = Some(XlsxReflection::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"relOff" => {
                    container.relative_offset = Some(XlsxRelativeOffset::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"softEdge" => {
                    container.soft_edge = Some(XlsxSoftEdge::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tint" => {
                    container.tint = Some(XlsxTint::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xfrm" => {
                    container.transform = Some(XlsxTransformEffect::load(e)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(container)
    }
}
