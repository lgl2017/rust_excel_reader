use std::collections::BTreeMap;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::raw::drawing::{
    effect::{effect_container::XlsxEffectDag, effect_list::XlsxEffectList},
    fill::{
        blip_fill::XlsxBlipFill, gradient_fill::XlsxGradientFill, group_fill::XlsxGroupFill,
        no_fill::XlsxNoFill, pattern_fill::XlsxPatternFill, solid_fill::XlsxSolidFill,
    },
    font::{
        complex_sript_font::XlsxComplexScriptFont, east_asian_font::XlsxEastAsianFont,
        latin_font::XlsxLatinFont,
    },
    line::outline::XlsxOutline,
};

use super::{
    highlight_color::XlsxHighlightColor,
    hyperlink_on_event::{XlsxHyperlinkOnClick, XlsxHyperlinkOnMouseOver},
    right_to_left::XlsxRightToLeft,
    symbol_font::XlsxSymbolFont,
    underline::XlsxUnderline,
    underline_fill::XlsxUnderlineFill,
    underline_fill_text::XlsxUnderlineFillText,
    underline_follow_text::XlsxUnderlineFollowsText,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.defaultrunproperties?view=openxml-3.0.1
/// This element contains all default run level text properties for the text runs within a containing paragraph.
/// These properties are to be used when overriding properties have not been defined within the rPr element
// tag: defRPr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDefaultTextRunProperties {
    // extLst (Extension List)	Not supported

    // Child Elements	Subclause
    // blipFill (Picture Fill)	§20.1.8.14
    pub blip_fill: Option<XlsxBlipFill>,

    // cs (Complex Script Font)	§21.1.2.3.1
    pub cs: Option<XlsxComplexScriptFont>,

    // ea (East Asian Font)	§21.1.2.3.3
    pub ea: Option<XlsxEastAsianFont>,

    // effectDag (Effect Container)	§20.1.8.25
    pub effect_dag: Option<XlsxEffectDag>,

    // effectLst (Effect Container)	§20.1.8.26
    pub effect_list: Option<XlsxEffectList>,

    // gradFill (Gradient Fill)	§20.1.8.33
    pub gradient_fill: Option<XlsxGradientFill>,

    // grpFill (Group Fill)	§20.1.8.35
    pub group_fill: Option<XlsxGroupFill>,

    // highlight (Highlight Color)	§21.1.2.3.4
    pub highlight: Option<XlsxHighlightColor>,

    // hlinkClick (Click Hyperlink)	§21.1.2.3.5
    pub hlink_click: Option<XlsxHyperlinkOnClick>,

    // hlinkMouseOver (Mouse-Over Hyperlink)	§21.1.2.3.6
    pub hlink_mouse_over: Option<XlsxHyperlinkOnMouseOver>,

    // latin (Latin Font)	§21.1.2.3.7
    pub latin: Option<XlsxLatinFont>,

    // ln (Outline)	§20.1.2.2.24
    pub outline: Option<XlsxOutline>,

    // noFill (No Fill)	§20.1.8.44
    pub no_fill: Option<XlsxNoFill>,

    // pattFill (Pattern Fill)	§20.1.8.47
    pub pattern_fill: Option<XlsxPatternFill>,

    // rtl (Right to Left Run)	§21.1.2.2.8
    pub rtl: Option<XlsxRightToLeft>,

    // solidFill (Solid Fill)	§20.1.8.54
    pub solid_fill: Option<XlsxSolidFill>,

    // sym (Symbol Font)	§21.1.2.3.10
    pub symbol_font: Option<XlsxSymbolFont>,

    // uFill (Underline Fill)	§21.1.2.3.12
    pub underline_fill: Option<XlsxUnderlineFill>,

    // uFillTx (Underline Fill Properties Follow Text)	§21.1.2.3.13
    pub underline_fill_text: Option<XlsxUnderlineFillText>,

    // uLn (Underline Stroke)	§21.1.2.3.14
    pub underline_stroke: Option<XlsxUnderline>,

    // uLnTx (Underline Follows Text)
    pub underline_follow_text: Option<XlsxUnderlineFollowsText>,

    // attributes: undocumented
    pub extended_attributes: Option<BTreeMap<String, String>>,
}

impl XlsxDefaultTextRunProperties {
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
                    defaults.blip_fill = Some(XlsxBlipFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cs" => {
                    defaults.cs = Some(XlsxComplexScriptFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ea" => {
                    defaults.ea = Some(XlsxEastAsianFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectDag" => {
                    defaults.effect_dag = Some(XlsxEffectDag::load_effect_dag(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectLst" => {
                    defaults.effect_list = Some(XlsxEffectList::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradFill" => {
                    defaults.gradient_fill = Some(XlsxGradientFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpFill" => {
                    defaults.group_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"highlight" => {
                    defaults.highlight = XlsxHighlightColor::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlinkClick" => {
                    defaults.hlink_click = Some(XlsxHyperlinkOnClick::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlinkMouseOver" => {
                    defaults.hlink_mouse_over = Some(XlsxHyperlinkOnMouseOver::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"latin" => {
                    defaults.latin = Some(XlsxLatinFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ln" => {
                    defaults.outline = Some(XlsxOutline::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"noFill" => {
                    defaults.no_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pattFill" => {
                    defaults.pattern_fill = Some(XlsxPatternFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rtl" => {
                    defaults.rtl = Some(XlsxRightToLeft::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                    defaults.solid_fill = XlsxSolidFill::load(reader, b"solidFill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sym" => {
                    defaults.symbol_font = Some(XlsxSymbolFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uFill" => {
                    defaults.underline_fill = XlsxUnderlineFill::load(reader, b"uFill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uFillTx" => {
                    defaults.underline_fill_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uLn" => {
                    defaults.underline_stroke = Some(XlsxUnderline::load(reader, e)?);
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
