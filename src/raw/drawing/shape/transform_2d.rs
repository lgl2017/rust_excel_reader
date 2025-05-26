use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_int},
    raw::drawing::st_types::STAngle,
};

use super::{extents::XlsxExtents, offset::XlsxOffset};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transform2d?view=openxml-3.0.1
///
/// This element represents 2-D transforms for ordinary shapes.
///
/// Example
/// ```
/// <a:xfrm>
///   <a:off x="3200400" y="1600200"/>
///   <a:ext cx="1705233" cy="679622"/>
/// </a:xfrm>
/// ```
///
/// xfrm
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTransform2D {
    // Child Elements
    // ext (Extents)
    pub extents: Option<XlsxExtents>,

    // off (Offset)
    pub offset: Option<XlsxOffset>,

    // Attributes
    /// Specifies a horizontal flip.
    /// When true, this attribute defines that the shape is flipped horizontally about the center of its bounding box.
    ///
    /// flipH (Horizontal Flip)
    pub horizontal_flip: Option<bool>,

    /// Specifies a vertical flip.
    /// When true, this attribute defines that the group is flipped vertically about the center of its bounding box.
    ///
    /// flipV (Vertical Flip)
    pub vertical_flip: Option<bool>,

    /// Specifies the rotation angle of the Graphic Frame.
    ///
    /// rot (Rotation)
    pub rotation: Option<STAngle>,
}

impl XlsxTransform2D {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut transform = Self {
            extents: None,
            offset: None,
            horizontal_flip: None,
            vertical_flip: None,
            rotation: None,
        };
        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"flipH" => {
                            transform.horizontal_flip = string_to_bool(&string_value);
                        }
                        b"flipV" => {
                            transform.vertical_flip = string_to_bool(&string_value);
                        }
                        b"rot" => {
                            transform.rotation = string_to_int(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ext" => {
                    transform.extents = Some(XlsxExtents::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"off" => {
                    transform.offset = Some(XlsxOffset::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"xfrm" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(transform)
    }
}
