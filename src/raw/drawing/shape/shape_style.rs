use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use crate::raw::drawing::{
    effect::effect_reference::EffectReference, fill::fill_reference::FillReference,
    font::font_reference::FontReference, line::line_reference::LineReference,
};

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
pub struct ShapeStyle {
    // Child Elements	Subclause
    // effectRef (Effect Reference)
    pub effect_reference: Option<EffectReference>,

    // fillRef (Fill Reference)	§20.1.4.2.10
    pub fill_reference: Option<FillReference>,

    // fontRef (Font Reference)	§20.1.4.1.17
    pub font_reference: Option<FontReference>,

    // lnRef (Line Reference)
    pub line_reference: Option<LineReference>,
}

impl ShapeStyle {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
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
                    shape_style.effect_reference = Some(EffectReference::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillRef" => {
                    shape_style.fill_reference = Some(FillReference::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fontRef" => {
                    shape_style.font_reference = Some(FontReference::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lnRef" => {
                    shape_style.line_reference = Some(LineReference::load(reader, e)?);
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
