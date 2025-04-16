use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use super::{
    alpha_bi_level::AlphaBiLevel, alpha_ceiling::AlphaCeiling, alpha_floor::AlphaFloor,
    alpha_inverse::AlphaInverse, alpha_modulation::AlphaModulation,
    alpha_modulation_fixed::AlphaModulationFixed, alpha_outset::AlphaOutset,
    alpha_replace::AlphaReplace, bi_level::BiLevel, blend::Blend, blur::Blur,
    color_change::ColorChange, color_replacement::ColorReplacement, duotone::Duotone,
    effect::Effect, fill::Fill, fill_overlay::FillOverlay, glow::Glow, gray_scale::GrayScale,
    hue_saturation_luminance::Hsl, inner_shadow::InnerShadow, luminance::Luminance,
    outer_shadow::OuterShadow, preset_shadow::PresetShadow, reflection::Reflection,
    relative_offset::RelativeOffset, soft_edge::SoftEdge, tint::Tint, transform::TransformEffect,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectdag?view=openxml-3.0.1
pub type EffectDag = EffectContainer;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectcontainer?view=openxml-3.0.1
/// A list of effects.
#[derive(Debug, Clone, PartialEq)]
pub struct EffectContainer {
    // Child Elements	Subclause

    // alphaBiLevel (Alpha Bi-Level Effect)	§20.1.8.1
    pub alpha_bi_level: Option<AlphaBiLevel>,

    // alphaCeiling (Alpha Ceiling Effect)	§20.1.8.2
    pub alpha_ceiling: Option<AlphaCeiling>,

    // alphaFloor (Alpha Floor Effect)	§20.1.8.3
    pub alpha_floor: Option<AlphaFloor>,

    // alphaInv (Alpha Inverse Effect)	§20.1.8.4
    pub alpha_inv: Option<AlphaInverse>,

    // alphaMod (Alpha Modulate Effect)	§20.1.8.5
    pub alpha_mod: Option<AlphaModulation>,

    // alphaModFix (Alpha Modulate Fixed Effect)	§20.1.8.6
    pub alpha_mod_fix: Option<AlphaModulationFixed>,

    // alphaOutset (Alpha Inset/Outset Effect)	§20.1.8.7
    pub alpha_outset: Option<AlphaOutset>,

    // alphaRepl (Alpha Replace Effect)	§20.1.8.8
    pub alpha_repl: Option<AlphaReplace>,

    // biLevel (Bi-Level (Black/White) Effect)	§20.1.8.11
    pub bi_level: Option<BiLevel>,

    // blend (Blend Effect)	§20.1.8.12
    pub blend: Option<Blend>,

    // blur (Blur Effect)	§20.1.8.15
    pub blur: Option<Blur>,

    // clrChange (Color Change Effect)	§20.1.8.16
    pub clr_change: Option<ColorChange>,

    // clrRepl (Solid Color Replacement)	§20.1.8.18
    pub clr_repl: Option<ColorReplacement>,

    // cont (Effect Container)	§20.1.8.20
    pub cont: Option<Box<EffectContainer>>,

    // duotone (Duotone Effect)	§20.1.8.23
    pub duotone: Option<Duotone>,

    // effect (Effect)	§20.1.8.24
    pub effect: Option<Effect>,

    // fill (Fill)	§20.1.8.28
    pub fill: Option<Fill>,

    // fillOverlay (Fill Overlay Effect)	§20.1.8.29
    pub fill_overlay: Option<Box<FillOverlay>>,

    // glow (Glow Effect)	§20.1.8.32
    pub glow: Option<Glow>,

    // grayscl (Gray Scale Effect)	§20.1.8.34
    pub grayscl: Option<GrayScale>,

    // hsl (Hue Saturation Luminance Effect)	§20.1.8.39
    pub hsl: Option<Hsl>,

    // innerShdw (Inner Shadow Effect)	§20.1.8.40
    pub innder_shadow: Option<InnerShadow>,

    // lum (Luminance Effect)	§20.1.8.42
    pub lum: Option<Luminance>,

    // outerShdw (Outer Shadow Effect)	§20.1.8.45
    pub outer_shadow: Option<OuterShadow>,

    // prstShdw (Preset Shadow)	§20.1.8.49
    pub preset_shadow: Option<PresetShadow>,

    // reflection (Reflection Effect)	§20.1.8.50
    pub reflection: Option<Reflection>,

    // relOff (Relative Offset Effect)	§20.1.8.51
    pub relative_offset: Option<RelativeOffset>,

    // softEdge (Soft Edge Effect)	§20.1.8.53
    pub soft_edge: Option<SoftEdge>,

    // tint (Tint Effect)
    pub tint: Option<Tint>,

    // xfrm (Transform Effect)
    pub transform: Option<TransformEffect>,

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

impl EffectContainer {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        return EffectContainer::load_helper(reader, e, b"cont");
    }

    pub(crate) fn load_effect_dag(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        return EffectContainer::load_helper(reader, e, b"effectDag");
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
                    container.alpha_bi_level = Some(AlphaBiLevel::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaCeiling" => {
                    container.alpha_ceiling = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaFloor" => {
                    container.alpha_floor = Some(true);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaInv" => {
                    container.alpha_inv = AlphaInverse::load(reader, b"alphaInv")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaMod" => {
                    container.alpha_mod = Some(AlphaModulation::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaModFix" => {
                    container.alpha_mod_fix = Some(AlphaModulationFixed::load(e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaOutset" => {
                    container.alpha_outset = Some(AlphaOutset::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaRepl" => {
                    container.alpha_repl = Some(AlphaReplace::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"biLevel" => {
                    container.bi_level = Some(BiLevel::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blend" => {
                    container.blend = Some(Blend::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blur" => {
                    container.blur = Some(Blur::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrChange" => {
                    container.clr_change = Some(ColorChange::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrRepl" => {
                    container.clr_repl = ColorReplacement::load(reader, b"clrRepl")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cont" => {
                    container.cont = Some(Box::new(EffectContainer::load(reader, e)?));
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"duotone" => {
                    container.duotone = Some(Duotone::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effect" => {
                    container.effect = Some(Effect::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                    container.fill = Fill::load(reader, b"fill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillOverlay" => {
                    if let Some(fill_overlay) = FillOverlay::load(reader, b"fillOverlay")? {
                        container.fill_overlay = Some(Box::new(fill_overlay));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"glow" => {
                    container.glow = Some(Glow::load(reader, e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grayscl" => {
                    container.grayscl = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hsl" => {
                    container.hsl = Some(Hsl::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"innerShdw" => {
                    container.innder_shadow = Some(InnerShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lum" => {
                    container.lum = Some(Luminance::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"outerShdw" => {
                    container.outer_shadow = Some(OuterShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"prstShdw" => {
                    container.preset_shadow = Some(PresetShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"reflection" => {
                    container.reflection = Some(Reflection::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"relOff" => {
                    container.relative_offset = Some(RelativeOffset::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"softEdge" => {
                    container.soft_edge = Some(SoftEdge::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tint" => {
                    container.tint = Some(Tint::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xfrm" => {
                    container.transform = Some(TransformEffect::load(e)?);
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
