use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::{
    excel::XmlReader,
    raw::drawing::shape::{
        end_connection::XlsxEndConnection, start_connection::XlsxStartConnection,
    },
};

use super::connection_shape_locks::{load_connection_shape_locks, XlsxConnectionShapeLocks};

/// - NonVisualConnectorShapeDrawingProperties: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualconnectorshapedrawingproperties?view=openxml-3.0.1
/// - Spreadsheet.NonVisualConnectorShapeDrawingProperties: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.nonvisualconnectorshapedrawingproperties?view=openxml-3.0.1
///
/// This element specifies the non-visual drawing properties for a connector shape.
/// These non-visual properties are properties that the generating application would utilize when rendering the slide surface.
///
/// cNvCxnSpPr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualConnectorShapeDrawingProperties {
    // Child Elements
    // extLst (Extension List)	Not supported
    /// cxnSpLocks (Connection Shape Locks)	§20.1.2.2.11
    pub connection_shape_locks: Option<XlsxConnectionShapeLocks>,

    /// endCxn (Connection End)	§20.1.2.2.13
    pub end_connection: Option<XlsxEndConnection>,

    /// stCxn (Connection Start)
    pub start_connection: Option<XlsxStartConnection>,
}

impl XlsxNonVisualConnectorShapeDrawingProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            connection_shape_locks: None,
            end_connection: None,
            start_connection: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cxnSpLocks" => {
                    properties.connection_shape_locks =
                        Some(load_connection_shape_locks(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"endCxn" => {
                    properties.end_connection = Some(XlsxEndConnection::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"stCxn" => {
                    properties.start_connection = Some(XlsxStartConnection::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cNvCxnSpPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at XlsxNonVisualConnectorShapeDrawingProperties: `cNvCxnSpPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
