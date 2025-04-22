use crate::excel::XmlReader;
use crate::raw::drawing::{
    effect::effect_reference::XlsxEffectReference, fill::fill_reference::XlsxFillReference,
    font::font_reference::XlsxFontReference, line::line_reference::XlsxLineReference,
};
use std::io::Read;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapestyle?view=openxml-3.0.1
///
/// This element specifies the style information for a shape.
///
/// Example
/// ```
/// <style>
///     <lnRef idx="1">
///         <schemeClr val="accent2"/>
///     </lnRef>
///     <fillRef idx="0">
///         <schemeClr val="accent2"/>
///     </fillRef>
///     <effectRef idx="0">
///         <schemeClr val="accent2"/>
///     </effectRef>
///     <fontRef idx="minor">
///         <schemeClr val="tx1"/>
///     </fontRef>
/// </style>
/// ```
// tag: style
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxShapeStyle {
    // Child Elements	Subclause
    // effectRef (Effect Reference)
    pub effect_reference: Option<XlsxEffectReference>,

    // fillRef (Fill Reference)	§20.1.4.2.10
    pub fill_reference: Option<XlsxFillReference>,

    // fontRef (Font Reference)	§20.1.4.1.17
    pub font_reference: Option<XlsxFontReference>,

    // lnRef (Line Reference)
    pub line_reference: Option<XlsxLineReference>,
}

impl XlsxShapeStyle {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut shape_style = Self {
            effect_reference: None,
            fill_reference: None,
            font_reference: None,
            line_reference: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectRef" => {
                    shape_style.effect_reference = Some(XlsxEffectReference::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillRef" => {
                    shape_style.fill_reference = Some(XlsxFillReference::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fontRef" => {
                    shape_style.font_reference = Some(XlsxFontReference::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lnRef" => {
                    shape_style.line_reference = Some(XlsxLineReference::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"style" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(shape_style)
    }
}
