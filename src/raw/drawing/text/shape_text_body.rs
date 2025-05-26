use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use super::paragraph::{text_list_style::XlsxTextListStyle, text_paragraphs::XlsxTextParagraphs};
use crate::{excel::XmlReader, raw::drawing::text::body_properties::XlsxBodyProperties};

/// - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textbody?view=openxml-3.0.1
/// - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.textbody?view=openxml-3.0.1
///
/// This element specifies the existence of text to be contained within the corresponding shape.
/// All visible text and visible text related properties are contained within this element.
/// There can be multiple paragraphs and within paragraphs multiple runs of text.
///
/// Example:
/// ```
/// <xdr:txBody>
/// <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip"
///     vert="horz" wrap="square" lIns="50800" tIns="50800" rIns="50800"
///     bIns="50800" numCol="1" spcCol="38100" rtlCol="0" anchor="t">
///     <a:spAutoFit />
/// </a:bodyPr>
/// <a:lstStyle />
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
/// <a:p>
///     <a:r>
///         <a:rPr lang="en-US" sz="1100" />
///         <a:t>text box</a:t>
///     </a:r>
/// </a:p>
/// <a:p>
///     <a:pPr marL="228600" indent="-228600">
///         <a:buFont typeface="+mj-lt" />
///         <a:buAutoNum type="arabicPeriod" />
///     </a:pPr>
///     <a:endParaRPr lang="en-US" sz="1100" />
/// </a:p>
/// </xdr:txBody>
/// ```
///
/// txBody (Shape Text Body)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxShapeTextBody {
    // Child Elements
    // bodyPr (Body Properties)
    pub body_properties: Option<XlsxBodyProperties>,

    // lstStyle (Text List Styles)
    pub text_list_style: Option<Box<XlsxTextListStyle>>,

    // p (Text Paragraphs)
    pub text_paragraph: Option<Vec<XlsxTextParagraphs>>,
}

impl XlsxShapeTextBody {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut text_body = Self {
            body_properties: None,
            text_list_style: None,
            text_paragraph: None,
        };

        let mut paragraphs: Vec<XlsxTextParagraphs> = vec![];

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bodyPr" => {
                    text_body.body_properties = Some(XlsxBodyProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lstStyle" => {
                    text_body.text_list_style = Some(Box::new(XlsxTextListStyle::load(reader)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"p" => {
                    paragraphs.push(XlsxTextParagraphs::load(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"txBody" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `txBody`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        text_body.text_paragraph = Some(paragraphs);

        return Ok(text_body);
    }
}
