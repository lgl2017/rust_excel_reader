use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::content_part_locks::{load_content_part_locks, XlsxContentPartLocks};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.drawing.nonvisualinkcontentpartproperties?view=openxml-3.0.1
///
/// cNvContentPartPr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualInkContentPartProperties {
    //  child elements:
    // DocumentFormat.OpenXml.Office2010.Drawing.OfficeArtExtensionList <a14:extLst>: Not supported

    // <a14:cpLocks>
    pub content_part_locks: Option<XlsxContentPartLocks>,
}

impl XlsxNonVisualInkContentPartProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            content_part_locks: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cpLocks" => {
                    properties.content_part_locks = Some(load_content_part_locks(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cNvContentPartPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at XlsxNonVisualInkContentPartProperties: `cNvContentPartPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
