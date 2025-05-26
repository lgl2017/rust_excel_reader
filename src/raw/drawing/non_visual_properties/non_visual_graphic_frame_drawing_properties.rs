use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::graphic_frame_locks::XlsxGraphicFrameLocks;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.nonvisualgraphicframedrawingproperties?view=openxml-3.0.1
///
/// This element specifies the non-visual properties for a single graphical object frame within a spreadsheet.
/// These are the set of properties of a frame which do not affect its display within a spreadsheet.
///
/// cNvGraphicFramePr (Non-Visual Graphic Frame Drawing Properties)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualGraphicFrameDrawingProperties {
    // Child Elements
    // extLst (Extension List): Not supported
    // graphicFrameLocks (Graphic Frame Locks)
    pub graphic_frame_locks: Option<XlsxGraphicFrameLocks>,
}

impl XlsxNonVisualGraphicFrameDrawingProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            graphic_frame_locks: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"graphicFrameLocks" => {
                    properties.graphic_frame_locks = Some(XlsxGraphicFrameLocks::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cNvGraphicFramePr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at `cNvGraphicFramePr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
