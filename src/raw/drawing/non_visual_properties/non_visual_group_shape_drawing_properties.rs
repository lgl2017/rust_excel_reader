use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::group_shape_locks::XlsxGroupShapeLocks;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualgroupshapedrawingproperties?view=openxml-3.0.1
///
/// This element specifies the non-visual drawing properties for a group shape.
/// These non-visual properties are properties that the generating application would utilize when rendering the slide surface.
///
/// Example:
/// ```
/// <p:nvGrpSpPr>
///     <p:cNvPr id="10" name="Group 9"/>
///     <p:cNvGrpSpPr/>
///     <p:nvPr/>
/// </p:nvGrpSpPr>
/// ```
///
/// cNvGrpSpPr (Non-Visual Group Shape Drawing Properties)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualGroupShapeDrawingProperties {
    // Child Elements
    // extLst (Extension List) Not supported

    // grpSpLocks (Group Shape Locks)
    pub group_shape_locks: Option<XlsxGroupShapeLocks>,
}

impl XlsxNonVisualGroupShapeDrawingProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            group_shape_locks: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"graphicFrameLocks" => {
                    properties.group_shape_locks = Some(XlsxGroupShapeLocks::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cNvGrpSpPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at `cNvGrpSpPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
