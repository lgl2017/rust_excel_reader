use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

use crate::{common_types::Text, excel::XmlReader, helper::extract_text_contents};

use super::run_properties::XlsxRunProperties;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.run?view=openxml-3.0.1
///
/// This element represents a run of rich text.
///
/// A rich text run is a region of text that share a common set of properties, such as formatting properties.
/// The properties are defined in the rPr element, and the text displayed to the user is defined in the Text (t) element.
///
/// Example:
/// ```
/// <si>
///     <r>
///         <rPr>
///             <sz val="10" />
///             <color indexed="8" />
///             <rFont val="Helvetica Neue" />
///         </rPr>
///         <t>123</t>
///     </r>
///     <r>
///         <rPr>
///             <b val="1" />
///             <sz val="10" />
///             <color indexed="8" />
///             <rFont val="Helvetica Neue" />
///         </rPr>
///         <t>4</t>
///     </r>
/// </si>
/// ```
/// r (Rich Text Run)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxRichTextRun {
    // Child Elements	Subclause
    // rPr (Run Properties)	§18.4.7
    pub run_properties: Option<XlsxRunProperties>,

    // t (Text)
    pub text: Option<Text>,
}

impl XlsxRichTextRun {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut run = Self {
            run_properties: None,
            text: None,
        };
        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => {
                    run.run_properties = Some(XlsxRunProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" => {
                    run.text = Some(extract_text_contents(reader, b"t")?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"r" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `r`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        return Ok(run);
    }
}
