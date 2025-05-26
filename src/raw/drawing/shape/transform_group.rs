use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_int},
};

use super::{
    extents::{XlsxChildExtent, XlsxExtents},
    offset::{XlsxChildOffset, XlsxOffset},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transformgroup?view=openxml-3.0.1
///
/// This element is nearly identical to the representation of 2-D transforms for ordinary shapes (§20.1.7.6).
/// The only addition is a member to represent the Child offset and the Child extents.
///
/// Example
/// ```
/// <a:xfrm>
///     <a:off x="838200" y="990600"/>
///     <a:ext cx="2426208" cy="978408"/>
///     <a:chOff x="838200" y="990600"/>
///     <a:chExt cx="2426208" cy="978408"/>
/// </a:xfrm>
/// ```
///
/// tag: xfrm (2D Transform for Grouped Objects)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTransformGroup {
    // Child Elements
    // chExt (Child Extents)
    pub child_extents: Option<XlsxChildExtent>,

    // chOff (Child Offset)
    pub child_offset: Option<XlsxChildOffset>,

    // ext (Extents)
    pub extents: Option<XlsxExtents>,

    // off (Offset)
    pub offset: Option<XlsxOffset>,

    // Attributes
    ///Specifies a horizontal flip.
    /// When true, this attribute defines that the shape is flipped horizontally about the center of its bounding box.
    // flipH (Horizontal Flip)
    pub horizontal_flip: Option<bool>,

    /// Specifies a vertical flip.
    /// When true, this attribute defines that the group is flipped vertically about the center of its bounding box.
    // flipV (Vertical Flip)
    pub vertical_flip: Option<bool>,

    /// Specifies the rotation angle of the Graphic Frame.
    // rot (Rotation)
    pub rotation: Option<i64>,
}

impl XlsxTransformGroup {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut transform = Self {
            child_extents: None,
            child_offset: None,
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"chExt" => {
                    transform.child_extents = Some(XlsxChildExtent::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"chOff" => {
                    transform.child_offset = Some(XlsxChildOffset::load(e)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"xfrm" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `xfrm`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(transform)
    }
}
