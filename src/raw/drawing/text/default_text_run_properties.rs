use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::excel::XmlReader;

use crate::helper::{string_to_bool, string_to_int, string_to_unsignedint};
use crate::raw::drawing::st_types::{STPercentage, STPositiveTextPoint, STTextPoint};
use crate::raw::drawing::{
    effect::{effect_container::XlsxEffectDag, effect_list::XlsxEffectList},
    fill::{
        blip_fill::XlsxBlipFill, gradient_fill::XlsxGradientFill, group_fill::XlsxGroupFill,
        no_fill::XlsxNoFill, pattern_fill::XlsxPatternFill, solid_fill::XlsxSolidFill,
    },
    line::outline::XlsxOutline,
};

use super::{
    font::{
        complex_sript_font::XlsxComplexScriptFont, east_asian_font::XlsxEastAsianFont,
        latin_font::XlsxLatinFont,
    },
    highlight_color::XlsxHighlightColor,
    hyperlink_on_event::{XlsxHyperlinkOnClick, XlsxHyperlinkOnMouseOver},
    right_to_left::XlsxRightToLeft,
    symbol_font::XlsxSymbolFont,
    underline::XlsxUnderline,
    underline_fill::XlsxUnderlineFill,
    underline_fill_text::XlsxUnderlineFillText,
    underline_follow_text::XlsxUnderlineFollowsText,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.runproperties?view=openxml-3.0.1
///
/// This element contains all run level text properties for the text runs within a containing paragraph.
///
/// Example:
/// ```
/// <a:r>
///     <a:rPr kumimoji="0" lang="en-US" sz="1100" b="0" i="0" u="none"
///         strike="noStrike" cap="none" spc="0" normalizeH="0" baseline="0">
///         <a:ln>
///             <a:noFill />
///         </a:ln>
///         <a:solidFill>
///             <a:srgbClr val="000000" />
///         </a:solidFill>
///         <a:effectLst />
///         <a:uFillTx />
///         <a:latin typeface="+mn-lt" />
///         <a:ea typeface="+mn-ea" />
///         <a:cs typeface="+mn-cs" />
///         <a:sym typeface="Helvetica Neue" />
///     </a:rPr>
///     <a:t>Text</a:t>
/// </a:r>
/// ```
///
/// rPr (Text Run Properties)
pub type XlsxTextRunProperties = XlsxDefaultTextRunProperties;

pub(crate) fn load_text_run_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxTextRunProperties> {
    return XlsxTextRunProperties::load_helper(reader, e, b"rPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.endparagraphrunproperties?view=openxml-3.0.1
///
/// This element specifies the text run properties that are to be used if another run is inserted after the last run specified.
/// This effectively saves the run property state so that it can be applied when the user enters additional text.
/// If this element is omitted, then the application can determine which default properties to apply.
/// It is recommended that this element be specified at the end of the list of text runs within the paragraph so that an orderly list is maintained.
///
/// endParaRPr (End Paragraph Run Properties)
pub type XlsxEndParagraphRunProperties = XlsxDefaultTextRunProperties;

pub(crate) fn load_end_paragraph_run_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxEndParagraphRunProperties> {
    return XlsxEndParagraphRunProperties::load_helper(reader, e, b"endParaRPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.defaultrunproperties?view=openxml-3.0.1
///
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

    // attributes:
    // Inherited from (TextCharacterPropertiesType)[https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textcharacterpropertiestype?view=openxml-3.0.1]
    /// AlternativeLanguage
    pub alternative_language: Option<String>,

    /// Baseline
    pub baseline: Option<STPercentage>,

    // Bold
    pub bold: Option<bool>,

    // Bookmark
    pub bookmark: Option<String>,

    /// Capital
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textcapsvalues?view=openxml-3.0.1
    pub capital: Option<String>,

    /// Dirty
    pub dirty: Option<bool>,

    /// FontSize
    pub font_size: Option<STPositiveTextPoint>,

    /// Italic
    pub italic: Option<bool>,

    /// Kerning
    pub kerning: Option<STPositiveTextPoint>,

    /// Kumimoji
    pub kumimoji: Option<bool>,

    /// Language
    pub language: Option<String>,

    /// NoProof
    pub no_proof: Option<bool>,

    /// NormalizeHeight
    pub normalize_height: Option<bool>,

    /// SmartTagClean
    pub smart_tag_clean: Option<bool>,

    /// SmartTagId
    pub smart_tag_id: Option<u64>,

    /// Spacing
    pub spacing: Option<STTextPoint>,

    /// SpellingError
    pub spelling_error: Option<bool>,

    /// Strike
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textcharacterpropertiestype.strike?view=openxml-3.0.1#documentformat-openxml-drawing-textcharacterpropertiestype-strike
    pub strike: Option<String>,

    /// Underline
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textunderlinevalues?view=openxml-3.0.1
    pub underline: Option<String>,
}

impl XlsxDefaultTextRunProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        return Self::load_helper(reader, e, b"defRPr");
    }

    fn load_helper(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
        tag: &[u8],
    ) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut properties = Self {
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
            // attributes:
            // Inherited from (TextCharacterPropertiesType)[https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textcharacterpropertiestype?view=openxml-3.0.1]
            alternative_language: None,
            baseline: None,
            bold: None,
            bookmark: None,
            capital: None,
            dirty: None,
            font_size: None,
            italic: None,
            kerning: None,
            kumimoji: None,
            language: None,
            no_proof: None,
            normalize_height: None,
            smart_tag_clean: None,
            smart_tag_id: None,
            spacing: None,
            spelling_error: None,
            strike: None,
            underline: None,
        };

        for a in e.attributes() {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"altLang" => {
                            properties.alternative_language = Some(string_value);
                        }
                        b"baseline" => {
                            properties.baseline = string_to_int(&string_value);
                        }
                        b"b" => {
                            properties.bold = string_to_bool(&string_value);
                        }
                        b"bmk" => {
                            properties.bookmark = Some(string_value);
                        }
                        b"cap" => {
                            properties.capital = Some(string_value);
                        }
                        b"dirty" => {
                            properties.dirty = string_to_bool(&string_value);
                        }
                        b"sz" => {
                            properties.font_size = string_to_unsignedint(&string_value);
                        }
                        b"i" => {
                            properties.italic = string_to_bool(&string_value);
                        }
                        b"kern" => {
                            properties.kerning = string_to_unsignedint(&string_value);
                        }
                        b"kumimoji" => {
                            properties.kumimoji = string_to_bool(&string_value);
                        }
                        b"lang" => {
                            properties.language = Some(string_value);
                        }
                        b"noProof" => {
                            properties.no_proof = string_to_bool(&string_value);
                        }
                        b"normalizeH" => {
                            properties.normalize_height = string_to_bool(&string_value);
                        }
                        b"smtClean" => {
                            properties.smart_tag_clean = string_to_bool(&string_value);
                        }
                        b"smtId" => {
                            properties.smart_tag_id = string_to_unsignedint(&string_value);
                        }
                        b"spc" => {
                            properties.spacing = string_to_unsignedint(&string_value);
                        }
                        b"err" => {
                            properties.spelling_error = string_to_bool(&string_value);
                        }
                        b"strike" => {
                            properties.strike = Some(string_value);
                        }
                        b"u" => {
                            properties.underline = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blipFill" => {
                    properties.blip_fill = Some(XlsxBlipFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cs" => {
                    properties.cs = Some(XlsxComplexScriptFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ea" => {
                    properties.ea = Some(XlsxEastAsianFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectDag" => {
                    properties.effect_dag = Some(XlsxEffectDag::load_effect_dag(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectLst" => {
                    properties.effect_list = Some(XlsxEffectList::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradFill" => {
                    properties.gradient_fill = Some(XlsxGradientFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpFill" => {
                    properties.group_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"highlight" => {
                    properties.highlight = XlsxHighlightColor::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlinkClick" => {
                    properties.hlink_click = Some(XlsxHyperlinkOnClick::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlinkMouseOver" => {
                    properties.hlink_mouse_over = Some(XlsxHyperlinkOnMouseOver::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"latin" => {
                    properties.latin = Some(XlsxLatinFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ln" => {
                    properties.outline = Some(XlsxOutline::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"noFill" => {
                    properties.no_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pattFill" => {
                    properties.pattern_fill = Some(XlsxPatternFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rtl" => {
                    properties.rtl = Some(XlsxRightToLeft::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                    properties.solid_fill = XlsxSolidFill::load(reader, b"solidFill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sym" => {
                    properties.symbol_font = Some(XlsxSymbolFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uFill" => {
                    properties.underline_fill = XlsxUnderlineFill::load(reader, b"uFill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uFillTx" => {
                    properties.underline_fill_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uLn" => {
                    properties.underline_stroke = Some(XlsxUnderline::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"uLnTx" => {
                    properties.underline_follow_text = Some(true);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(properties)
    }
}
