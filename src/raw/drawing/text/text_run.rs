use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use super::default_text_run_properties::{load_text_run_properties, XlsxTextRunProperties};
use crate::{common_types::Text, excel::XmlReader, helper::extract_text_contents};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.run?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:r>
///     <a:rPr kumimoji="0" lang="en-US" sz="1100" b="0" i="0" u="none"
///         strike="noStrike" cap="none" spc="0" normalizeH="0" baseline="0">
///         <a:ln>
///             <a:noFill />
///         </a:ln>
///         <a:solidFill>
///             <a:srgbClr val="000000" />
///         </a:solidFill>
///         <a:effectLst />
///         <a:uFillTx />
///         <a:latin typeface="+mn-lt" />
///         <a:ea typeface="+mn-ea" />
///         <a:cs typeface="+mn-cs" />
///         <a:sym typeface="Helvetica Neue" />
///     </a:rPr>
///     <a:t>Text</a:t>
/// </a:r>
/// ```
///
/// r (Text Run)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTextRun {
    // Child Elements

    // rPr (Text Run Properties)	§21.1.2.3.9
    pub run_properties: Option<XlsxTextRunProperties>,

    // t (Text String)
    pub text: Option<Text>,
}

impl XlsxTextRun {
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
                    run.run_properties = Some(load_text_run_properties(reader, e)?);
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
