use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use super::{
    paragraph::text_list_style::ListStyle,
    shape::{shape_properties::ShapeProperties, shape_style::ShapeStyle},
    text::body_properties::BodyProperties,
};

pub mod line_default;
pub mod object_defaults;
pub mod shape_default;
pub mod text_default;

#[derive(Debug, Clone, PartialEq)]
pub struct DefaultBase {
    // extLst (Extension List) Not supported
    // Child Elements
    // bodyPr (Body Properties)	§21.1.2.1.1
    pub body_properties: Option<BodyProperties>,

    // lstStyle (Text List Styles)	§21.1.2.4.12
    pub text_list_style: Option<Box<ListStyle>>,

    // spPr (Shape Properties)	§20.1.2.2.35
    pub shape_properties: Option<ShapeProperties>,

    // style (Shape Style)
    pub style: Option<ShapeStyle>,
}

impl DefaultBase {
    pub(crate) fn load(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut defaults = Self {
            body_properties: None,
            text_list_style: None,
            shape_properties: None,
            style: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bodyPr" => {
                    defaults.body_properties = Some(BodyProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lstStyle" => {
                    defaults.text_list_style = Some(Box::new(ListStyle::load(reader)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                    defaults.shape_properties = Some(ShapeProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"style" => {
                    defaults.style = Some(ShapeStyle::load(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(defaults)
    }
}
