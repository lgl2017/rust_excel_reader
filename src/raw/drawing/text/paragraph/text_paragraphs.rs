use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::{
    excel::XmlReader,
    raw::drawing::text::{text_field::XlsxTextField, text_run::XlsxTextRun},
};

use super::{
    super::default_text_run_properties::{
        load_end_paragraph_run_properties, XlsxEndParagraphRunProperties,
    },
    line_break::XlsxTextLineBreak,
    paragraph_properties::{load_text_paragraph_properties, XlsxTextParagraphProperties},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraph?view=openxml-3.0.1
///
/// This element specifies the presence of a paragraph of text within the containing text body.
/// The paragraph is the highest level text separation mechanism within a text body.
/// A paragraph can contain text paragraph properties associated with the paragraph.
/// If no properties are listed then properties specified in the defPPr element are used.
///
/// Example:
/// ```
/// <a:p>
///     <a:pPr marL="0" marR="0" indent="0" algn="l" defTabSz="457200" rtl="0"
///         fontAlgn="auto" latinLnBrk="0" hangingPunct="0">
///         <a:lnSpc>
///             <a:spcPct val="100000" />
///         </a:lnSpc>
///         <a:spcBef>
///             <a:spcPts val="0" />
///         </a:spcBef>
///         <a:spcAft>
///             <a:spcPts val="0" />
///         </a:spcAft>
///         <a:buClrTx />
///         <a:buSzTx />
///         <a:buFontTx />
///         <a:buNone />
///     </a:pPr>
///     <a:r>
///         <a:rPr kumimoji="0" lang="en-US" sz="1100" b="0" i="0" u="none"
///             strike="noStrike" cap="none" spc="0" normalizeH="0" baseline="0">
///             <a:ln>
///                 <a:noFill />
///             </a:ln>
///             <a:solidFill>
///                 <a:srgbClr val="000000" />
///             </a:solidFill>
///             <a:effectLst />
///             <a:uFillTx />
///             <a:latin typeface="+mn-lt" />
///             <a:ea typeface="+mn-ea" />
///             <a:cs typeface="+mn-cs" />
///             <a:sym typeface="Helvetica Neue" />
///         </a:rPr>
///         <a:t>Text</a:t>
///     </a:r>
/// </a:p>
/// ```
///
/// p (Text Paragraphs)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTextParagraphs {
    // Child Elements

    // endParaRPr (End Paragraph Run Properties)
    pub end_paragraph_run_properties: Option<XlsxEndParagraphRunProperties>,

    // pPr (Text Paragraph Properties)
    pub paragraph_properties: Option<Box<XlsxTextParagraphProperties>>,

    pub runs: Option<Vec<XlsxRunType>>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum XlsxRunType {
    // r (Text Run)
    Text(XlsxTextRun),

    // br (Text Line Break)
    LineBreak(XlsxTextLineBreak),

    // fld (Text Field)
    TextField(XlsxTextField),
}

impl XlsxTextParagraphs {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut field = Self {
            paragraph_properties: None,
            end_paragraph_run_properties: None,
            runs: None,
        };
        let mut runs: Vec<XlsxRunType> = vec![];

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"br" => {
                    let br = XlsxTextLineBreak::load(reader)?;
                    runs.push(XlsxRunType::LineBreak(br));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"endParaRPr" => {
                    field.end_paragraph_run_properties =
                        Some(load_end_paragraph_run_properties(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fld" => {
                    runs.push(XlsxRunType::TextField(XlsxTextField::load(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pPr" => {
                    field.paragraph_properties =
                        Some(Box::new(load_text_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"r" => {
                    runs.push(XlsxRunType::Text(XlsxTextRun::load(reader)?));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"p" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `p`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        field.runs = Some(runs);
        return Ok(field);
    }
}
