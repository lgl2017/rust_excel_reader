use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::string_to_bool,
    raw::drawing::{
        fill::blip_fill::XlsxBlipFill,
        non_visual_properties::non_visual_picture_properties::XlsxNonVisualPictureProperties,
        shape::{shape_properties::XlsxShapeProperties, shape_style::XlsxShapeStyle},
    },
};

/// - Picture: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.picture?view=openxml-3.0.1
/// - SpreadSheet.Picture: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.picture?view=openxml-3.0.1
///
/// This element specifies the existence of a picture object within the document.
///
/// Example:
/// ```
/// <p:pic>
///   <p:nvPicPr>
///     <p:cNvPr id="4" name="lake.JPG" descr="Picture of a Lake" />
///     <p:cNvPicPr>
///       <a:picLocks noChangeAspect="1"/>
///     </p:cNvPicPr>
///     <p:nvPr/>
///   </p:nvPicPr>
///   <p:blipFill>
///   …  </p:blipFill>
///   <p:spPr>
///   …  </p:spPr>
/// </p:pic>
/// ```
///
/// pic (Picture)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPicture {
    //Child Elements
    // extLst (Extension List) Not supported

    // blipFill (Picture Fill)
    pub blip_fill: Option<XlsxBlipFill>,

    // nvPicPr (Non-Visual Properties for a Picture)
    pub non_visual_picture_properties: Option<XlsxNonVisualPictureProperties>,

    // spPr (Shape Properties)
    pub shape_properties: Option<XlsxShapeProperties>,

    // style (Shape Style)
    pub shape_style: Option<XlsxShapeStyle>,

    // attributes for SpreadSheet.Picture
    /// Macro
    ///
    /// Reference to Custom Function
    pub r#macro: Option<String>,

    /// Published
    ///
    /// Publish to Server Flag
    pub published: Option<bool>,
}

impl XlsxPicture {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut picture = Self {
            blip_fill: None,
            non_visual_picture_properties: None,
            shape_properties: None,
            shape_style: None,
            r#macro: None,
            published: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"macro" => {
                            picture.r#macro = Some(string_value);
                            break;
                        }
                        b"fPublished" => {
                            picture.published = string_to_bool(&string_value);
                        }

                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blipFill" => {
                    picture.blip_fill = Some(XlsxBlipFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"nvPicPr" => {
                    picture.non_visual_picture_properties =
                        Some(XlsxNonVisualPictureProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                    picture.shape_properties = Some(XlsxShapeProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"style" => {
                    picture.shape_style = Some(XlsxShapeStyle::load(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pic" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at Picture: `pic`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(picture)
    }
}
