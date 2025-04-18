use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

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
/// A list of effects.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxEffectContainer {
    // Child Elements	Subclause

    // alphaBiLevel (Alpha Bi-Level Effect)	§20.1.8.1
    pub alpha_bi_level: Option<XlsxAlphaBiLevel>,

    // alphaCeiling (Alpha Ceiling Effect)	§20.1.8.2
    pub alpha_ceiling: Option<XlsxAlphaCeiling>,

    // alphaFloor (Alpha Floor Effect)	§20.1.8.3
    pub alpha_floor: Option<XlsxAlphaFloor>,

    // alphaInv (Alpha Inverse Effect)	§20.1.8.4
    pub alpha_inv: Option<XlsxAlphaInverse>,

    // alphaMod (Alpha Modulate Effect)	§20.1.8.5
    pub alpha_mod: Option<XlsxAlphaModulation>,

    // alphaModFix (Alpha Modulate Fixed Effect)	§20.1.8.6
    pub alpha_mod_fix: Option<XlsxAlphaModulationFixed>,

    // alphaOutset (Alpha Inset/Outset Effect)	§20.1.8.7
    pub alpha_outset: Option<XlsxAlphaOutset>,

    // alphaRepl (Alpha Replace Effect)	§20.1.8.8
    pub alpha_repl: Option<XlsxAlphaReplace>,

    // biLevel (Bi-Level (Black/White) Effect)	§20.1.8.11
    pub bi_level: Option<XlsxBiLevel>,

    // blend (Blend Effect)	§20.1.8.12
    pub blend: Option<XlsxBlend>,

    // blur (Blur Effect)	§20.1.8.15
    pub blur: Option<XlsxBlur>,

    // clrChange (Color Change Effect)	§20.1.8.16
    pub clr_change: Option<XlsxColorChange>,

    // clrRepl (Solid Color Replacement)	§20.1.8.18
    pub clr_repl: Option<XlsxColorReplacement>,

    // cont (Effect Container)	§20.1.8.20
    pub cont: Option<Box<XlsxEffectContainer>>,

    // duotone (Duotone Effect)	§20.1.8.23
    pub duotone: Option<XlsxDuotone>,

    // effect (Effect)	§20.1.8.24
    pub effect: Option<XlsxEffect>,

    // fill (Fill)	§20.1.8.28
    pub fill: Option<XlsxFill>,

    // fillOverlay (Fill Overlay Effect)	§20.1.8.29
    pub fill_overlay: Option<Box<XlsxFillOverlay>>,

    // glow (Glow Effect)	§20.1.8.32
    pub glow: Option<XlsxGlow>,

    // grayscl (Gray Scale Effect)	§20.1.8.34
    pub grayscl: Option<XlsxGrayScale>,

    // hsl (Hue Saturation Luminance Effect)	§20.1.8.39
    pub hsl: Option<XlsxHsl>,

    // innerShdw (Inner Shadow Effect)	§20.1.8.40
    pub innder_shadow: Option<XlsxInnerShadow>,

    // lum (Luminance Effect)	§20.1.8.42
    pub lum: Option<XlsxLuminance>,

    // outerShdw (Outer Shadow Effect)	§20.1.8.45
    pub outer_shadow: Option<XlsxOuterShadow>,

    // prstShdw (Preset Shadow)	§20.1.8.49
    pub preset_shadow: Option<XlsxPresetShadow>,

    // reflection (Reflection Effect)	§20.1.8.50
    pub reflection: Option<XlsxReflection>,

    // relOff (Relative Offset Effect)	§20.1.8.51
    pub relative_offset: Option<XlsxRelativeOffset>,

    // softEdge (Soft Edge Effect)	§20.1.8.53
    pub soft_edge: Option<XlsxSoftEdge>,

    // tint (Tint Effect)
    pub tint: Option<XlsxTint>,

    // xfrm (Transform Effect)
    pub transform: Option<XlsxTransformEffect>,

    // attributes
    /// Specifies an optional name for this list of effects, so that it can be referred to later.
    /// Shall be unique across all effect trees and effect containers.
    // tag: name (Name)
    pub name: Option<String>,

    /// type (Effect Container Type)
    /// Specifies the kind of container, either sibling or tree.
    /// allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainervalues?view=openxml-3.0.1
    pub r#type: Option<String>,
}

impl XlsxEffectContainer {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        return XlsxEffectContainer::load_helper(reader, e, b"cont");
    }

    pub(crate) fn load_effect_dag(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        return XlsxEffectContainer::load_helper(reader, e, b"effectDag");
    }

    fn load_helper(reader: &mut XmlReader, e: &BytesStart, tag: &[u8]) -> anyhow::Result<Self> {
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
                    container.duotone = Some(XlsxDuotone::load(reader)?);
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
