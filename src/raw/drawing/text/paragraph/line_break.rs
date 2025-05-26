use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::super::default_text_run_properties::{load_text_run_properties, XlsxTextRunProperties};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.break?view=openxml-3.0.1
///
/// This element specifies the existence of a vertical line break between two runs of text within a paragraph.
/// In addition to specifying a vertical space between two runs of text, this element can also have run properties specified via the rPr child element.
/// This sets the formatting of text for the line break so that if text is later inserted there that a new run can be generated with the correct formatting.
///
/// Example:
/// ```
/// <a:br>
/// <a:rPr lang="en-US" sz="1100">
///     <a:ln>
///         <a:solidFill>
///             <a:schemeClr val="accent5">
///                 <a:alpha val="99339" />
///             </a:schemeClr>
///         </a:solidFill>
///     </a:ln>
/// </a:rPr>
/// </a:br>
/// ```
/// br (Text Line Break)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTextLineBreak {
    // Child Elements
    // rPr (Text Run Properties)
    pub text_run_properties: Option<XlsxTextRunProperties>,
}

impl XlsxTextLineBreak {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut run = Self {
            text_run_properties: None,
        };
        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => {
                    run.text_run_properties = Some(load_text_run_properties(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"br" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `br`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        return Ok(run);
    }
}
