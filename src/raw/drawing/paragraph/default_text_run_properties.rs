use std::collections::BTreeMap;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::raw::drawing::{
    effect::{effect_container::EffectDag, effect_list::EffectList},
    fill::{
        blip_fill::BlipFill, gradient_fill::GradientFill, group_fill::GroupFill, no_fill::NoFill,
        pattern_fill::PatternFill, solid_fill::SolidFill,
    },
    font::{
        complex_sript_font::ComplexScriptFont, east_asian_font::EastAsianFont,
        latin_font::LatinFont,
    },
    line::outline::Outline,
};

use super::{
    highlight_color::HighlightColor,
    hyperlink_on_event::{HyperlinkOnClick, HyperlinkOnMouseOver},
    right_to_left::RightToLeft,
    symbol_font::SymbolFont,
    underline::Underline,
    underline_fill::UnderlineFill,
    underline_fill_text::UnderlineFillText,
    underline_follow_text::UnderlineFollowsText,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.defaultrunproperties?view=openxml-3.0.1
/// This element contains all default run level text properties for the text runs within a containing paragraph.
/// These properties are to be used when overriding properties have not been defined within the rPr element
// tag: defRPr
#[derive(Debug, Clone, PartialEq)]
pub struct DefaultTextRunProperties {
    // extLst (Extension List)	Not supported

    // Child Elements	Subclause
    // blipFill (Picture Fill)	§20.1.8.14
    pub blip_fill: Option<BlipFill>,

    // cs (Complex Script Font)	§21.1.2.3.1
    pub cs: Option<ComplexScriptFont>,

    // ea (East Asian Font)	§21.1.2.3.3
    pub ea: Option<EastAsianFont>,

    // effectDag (Effect Container)	§20.1.8.25
    pub effect_dag: Option<EffectDag>,

    // effectLst (Effect Container)	§20.1.8.26
    pub effect_list: Option<EffectList>,

    // gradFill (Gradient Fill)	§20.1.8.33
    pub gradient_fill: Option<GradientFill>,

    // grpFill (Group Fill)	§20.1.8.35
    pub group_fill: Option<GroupFill>,

    // highlight (Highlight Color)	§21.1.2.3.4
    pub highlight: Option<HighlightColor>,

    // hlinkClick (Click Hyperlink)	§21.1.2.3.5
    pub hlink_click: Option<HyperlinkOnClick>,

    // hlinkMouseOver (Mouse-Over Hyperlink)	§21.1.2.3.6
    pub hlink_mouse_over: Option<HyperlinkOnMouseOver>,

    // latin (Latin Font)	§21.1.2.3.7
    pub latin: Option<LatinFont>,

    // ln (Outline)	§20.1.2.2.24
    pub outline: Option<Outline>,

    // noFill (No Fill)	§20.1.8.44
    pub no_fill: Option<NoFill>,

    // pattFill (Pattern Fill)	§20.1.8.47
    pub pattern_fill: Option<PatternFill>,

    // rtl (Right to Left Run)	§21.1.2.2.8
    pub rtl: Option<RightToLeft>,

    // solidFill (Solid Fill)	§20.1.8.54
    pub solid_fill: Option<SolidFill>,

    // sym (Symbol Font)	§21.1.2.3.10
    pub symbol_font: Option<SymbolFont>,

    // uFill (Underline Fill)	§21.1.2.3.12
    pub underline_fill: Option<UnderlineFill>,

    // uFillTx (Underline Fill Properties Follow Text)	§21.1.2.3.13
    pub underline_fill_text: Option<UnderlineFillText>,

    // uLn (Underline Stroke)	§21.1.2.3.14
    pub underline_stroke: Option<Underline>,

    // uLnTx (Underline Follows Text)
    pub underline_follow_text: Option<UnderlineFollowsText>,

    // attributes: undocumented
    pub extended_attributes: Option<BTreeMap<String, String>>,
}

impl DefaultTextRunProperties {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut defaults = Self {
            blip_fill: None,
            cs: None,
            ea: None,
            effect_dag: None,
            effect_list: None,
            gradient_fill: None,
            group_fill: None,
            highlight: None,
            hlink_click: None,
            hlink_mouse_over: None,
            latin: None,
            outline: None,
            no_fill: None,
            pattern_fill: None,
            rtl: None,
            solid_fill: None,
            symbol_font: None,
            underline_fill: None,
            underline_fill_text: None,
            underline_stroke: None,
            underline_follow_text: None,
            // undocumented attributes
            extended_attributes: None,
        };

        let mut attributes: BTreeMap<String, String> = BTreeMap::new();

        for a in e.attributes() {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    let key = String::from_utf8(a.key.local_name().as_ref().to_vec())?;
                    attributes.insert(key, string_value);
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        defaults.extended_attributes = Some(attributes);

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blipFill" => {
                    defaults.blip_fill = Some(BlipFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cs" => {
                    defaults.cs = Some(ComplexScriptFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ea" => {
                    defaults.ea = Some(EastAsianFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectDag" => {
                    defaults.effect_dag = Some(EffectDag::load_effect_dag(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectLst" => {
                    defaults.effect_list = Some(EffectList::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradFill" => {
                    defaults.gradient_fill = Some(GradientFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpFill" => {
                    defaults.group_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"highlight" => {
                    defaults.highlight = HighlightColor::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlinkClick" => {
                    defaults.hlink_click = Some(HyperlinkOnClick::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlinkMouseOver" => {
                    defaults.hlink_mouse_over = Some(HyperlinkOnMouseOver::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"latin" => {
                    defaults.latin = Some(LatinFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ln" => {
                    defaults.outline = Some(Outline::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"noFill" => {
                    defaults.no_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pattFill" => {
                    defaults.pattern_fill = Some(PatternFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rtl" => {
                    defaults.rtl = Some(RightToLeft::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                    defaults.solid_fill = SolidFill::load(reader, b"solidFill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sym" => {
                    defaults.symbol_font = Some(SymbolFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uFill" => {
                    defaults.underline_fill = UnderlineFill::load(reader, b"uFill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uFillTx" => {
                    defaults.underline_fill_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uLn" => {
                    defaults.underline_stroke = Some(Underline::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uLnTx" => {
                    defaults.underline_follow_text = Some(true);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"defRPr" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(defaults)
    }
}
