use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use super::{
    blur::Blur, fill_overlay::FillOverlay, glow::Glow, inner_shadow::InnerShadow,
    outer_shadow::OuterShadow, preset_shadow::PresetShadow, reflection::Reflection,
    soft_edge::SoftEdge,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectlist?view=openxml-3.0.1
/// tag: effectLst
#[derive(Debug, Clone, PartialEq)]
pub struct EffectList {
    /// Child Elements

    /// blur (Blur Effect)	§20.1.8.15
    pub blur: Option<Blur>,

    /// fillOverlay (Fill Overlay Effect)	§20.1.8.29
    pub fill_overlay: Option<Box<FillOverlay>>,
    /// glow (Glow Effect)	§20.1.8.32
    pub glow: Option<Glow>,

    /// innerShdw (Inner Shadow Effect)	§20.1.8.40
    pub innder_shadow: Option<InnerShadow>,

    /// outerShdw (Outer Shadow Effect)	§20.1.8.45
    pub outer_shadow: Option<OuterShadow>,

    /// prstShdw (Preset Shadow)	§20.1.8.49
    pub preset_shadow: Option<PresetShadow>,

    /// reflection (Reflection Effect)	§20.1.8.50
    pub reflection: Option<Reflection>,

    /// softEdge (Soft Edge Effect)
    pub soft_edge: Option<SoftEdge>,
}

impl EffectList {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut list = Self {
            blur: None,
            fill_overlay: None,
            glow: None,
            innder_shadow: None,
            outer_shadow: None,
            preset_shadow: None,
            reflection: None,
            soft_edge: None,
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blur" => {
                    list.blur = Some(Blur::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillOverlay" => {
                    if let Some(fill_overlay) = FillOverlay::load(reader, b"fillOverlay")? {
                        list.fill_overlay = Some(Box::new(fill_overlay));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"glow" => {
                    list.glow = Some(Glow::load(reader, e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"innerShdw" => {
                    list.innder_shadow = Some(InnerShadow::load(reader, e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"outerShdw" => {
                    list.outer_shadow = Some(OuterShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"prstShdw" => {
                    list.preset_shadow = Some(PresetShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"reflection" => {
                    list.reflection = Some(Reflection::load(e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"softEdge" => {
                    list.soft_edge = Some(SoftEdge::load(e)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"effectLst" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(list)
    }
}
