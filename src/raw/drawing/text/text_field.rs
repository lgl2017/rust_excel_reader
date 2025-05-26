use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use super::default_text_run_properties::{load_text_run_properties, XlsxTextRunProperties};
use super::paragraph::paragraph_properties::{
    load_text_paragraph_properties, XlsxTextParagraphProperties,
};

use crate::{common_types::Text, excel::XmlReader, helper::extract_text_contents};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.field?view=openxml-3.0.1
///
/// This element specifies a text field which contains generated text that the application should update periodically.
/// Each piece of text when it is generated is given a unique identification number that is used to refer to a specific field.
/// At the time of creation the text field indicates the kind of text that should be used to update this field.
/// This update type is used so that all applications that did not create this text field can still know what kind of text it should be updated with. Thus the new application can then attach an update type to the text field id for continual updating.
///
/// Example:
/// ```
/// <a:fld id="{424CEEAC-8F67-4238-9622-1B74DC6E8318}" type="slidenum">
///     <a:rPr lang="en-US" smtClean="0"/>
///     <a:pPr/>
///     <a:t>3</a:t>
/// </a:fld>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTextField {
    // Child Elements
    // pPr (Text Paragraph Properties)
    pub paragraph_properties: Option<Box<XlsxTextParagraphProperties>>,

    // rPr (Text Run Properties)
    pub run_properties: Option<XlsxTextRunProperties>,

    // t (Text String)
    pub text: Option<Text>,

    /// Specifies the unique to this document, host specified token that is used to identify the field.
    ///
    /// This token is generated when the text field is created and persists in the file as the same token until the text field is removed.
    /// Any application should check the document for conflicting tokens before assigning a new token to a text field.
    pub id: Option<String>,

    /// Specifies the type of text that should be used to update this text field.
    ///
    /// This is used to inform the rendering application what text it should use to update this text field.
    /// There are no specific syntax restrictions placed on this attribute.
    /// The generating application can use it to represent any text that should be updated before rendering the presentation.
    ///
    /// For type `TxLink`, caculation (cell) reference is defined in the sp (Shpae) textlink property.
    pub r#type: Option<String>,
}

impl XlsxTextField {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut field = Self {
            run_properties: None,
            text: None,
            paragraph_properties: None,
            id: None,
            r#type: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"id" => {
                            field.id = Some(string_value);
                        }
                        b"type" => {
                            field.r#type = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pPr" => {
                    field.paragraph_properties =
                        Some(Box::new(load_text_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => {
                    field.run_properties = Some(load_text_run_properties(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" => {
                    field.text = Some(extract_text_contents(reader, b"t")?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fld" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `fld`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        return Ok(field);
    }
}
