use super::{
    blur::XlsxBlur, fill_overlay::XlsxFillOverlay, glow::XlsxGlow, inner_shadow::XlsxInnerShadow,
    outer_shadow::XlsxOuterShadow, preset_shadow::XlsxPresetShadow, reflection::XlsxReflection,
    soft_edge::XlsxSoftEdge,
};
use crate::excel::XmlReader;
use std::io::Read;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectlist?view=openxml-3.0.1
/// tag: effectLst
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxEffectList {
    /// Child Elements

    /// blur (Blur Effect)	§20.1.8.15
    pub blur: Option<XlsxBlur>,

    /// fillOverlay (Fill Overlay Effect)	§20.1.8.29
    pub fill_overlay: Option<Box<XlsxFillOverlay>>,
    /// glow (Glow Effect)	§20.1.8.32
    pub glow: Option<XlsxGlow>,

    /// innerShdw (Inner Shadow Effect)	§20.1.8.40
    pub innder_shadow: Option<XlsxInnerShadow>,

    /// outerShdw (Outer Shadow Effect)	§20.1.8.45
    pub outer_shadow: Option<XlsxOuterShadow>,

    /// prstShdw (Preset Shadow)	§20.1.8.49
    pub preset_shadow: Option<XlsxPresetShadow>,

    /// reflection (Reflection Effect)	§20.1.8.50
    pub reflection: Option<XlsxReflection>,

    /// softEdge (Soft Edge Effect)
    pub soft_edge: Option<XlsxSoftEdge>,
}

impl XlsxEffectList {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
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
                    list.blur = Some(XlsxBlur::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillOverlay" => {
                    if let Some(fill_overlay) = XlsxFillOverlay::load(reader, b"fillOverlay")? {
                        list.fill_overlay = Some(Box::new(fill_overlay));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"glow" => {
                    list.glow = Some(XlsxGlow::load(reader, e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"innerShdw" => {
                    list.innder_shadow = Some(XlsxInnerShadow::load(reader, e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"outerShdw" => {
                    list.outer_shadow = Some(XlsxOuterShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"prstShdw" => {
                    list.preset_shadow = Some(XlsxPresetShadow::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"reflection" => {
                    list.reflection = Some(XlsxReflection::load(e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"softEdge" => {
                    list.soft_edge = Some(XlsxSoftEdge::load(e)?);
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
