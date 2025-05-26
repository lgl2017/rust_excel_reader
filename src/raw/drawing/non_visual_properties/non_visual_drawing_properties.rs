use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_unsignedint},
    raw::drawing::text::hyperlink_on_event::{
        XlsxHyperlinkOnClick, XlsxHyperlinkOnHover, XlsxHyperlinkOnMouseOver,
    },
};

/// - NonVisualDrawingProperties: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.drawing.nonvisualdrawingproperties?view=openxml-3.0.1
/// - SpreadSheet.NonVisualDrawingProperties: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.nonvisualdrawingproperties?view=openxml-3.0.1
///
/// xdr14:cNvPr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualDrawingProperties {
    // children
    // extLst (Not Supported)
    /// hlinkClick (HyperlinkOnClick)
    ///
    /// Hyperlink associated with clicking or selecting the element.
    pub hlink_click: Option<XlsxHyperlinkOnClick>,

    /// hlinkHover (HyperlinkOnHover)
    ///
    /// Hyperlink associated with hovering over the element.
    pub hlink_hover: Option<XlsxHyperlinkOnHover>,

    // attributes
    /// descr (Description)
    ///
    /// Description of the drawing element.
    pub description: Option<String>,

    /// hidden (Hidden)
    ///
    /// Flag determining to show or hide this element.
    pub hidden: Option<bool>,

    /// id (Id)
    ///
    /// Application defined unique identifier.
    pub id: Option<u64>,

    /// name (Name)
    ///
    /// Name compatible with Object Model (non-unique).
    pub name: Option<String>,

    /// title (Title)
    pub title: Option<String>,
}

impl XlsxNonVisualDrawingProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            hlink_click: None,
            hlink_hover: None,
            description: None,
            hidden: None,
            id: None,
            name: None,
            title: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"descr" => {
                            properties.description = Some(string_value);
                        }
                        b"hidden" => {
                            properties.hidden = string_to_bool(&string_value);
                        }
                        b"id" => {
                            properties.id = string_to_unsignedint(&string_value);
                        }
                        b"name" => {
                            properties.name = Some(string_value);
                        }
                        b"title" => {
                            properties.title = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlinkClick" => {
                    properties.hlink_click = Some(XlsxHyperlinkOnClick::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlinkMouseOver" => {
                    properties.hlink_hover = Some(XlsxHyperlinkOnMouseOver::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cNvPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at XlsxNonVisualDrawingProperties: `cNvPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
