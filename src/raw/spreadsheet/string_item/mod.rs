pub mod phonetic_properties;
pub mod phonetic_run;
pub mod rich_text_run;
pub mod run_properties;
pub mod text;

use anyhow::bail;
use phonetic_properties::XlsxPhoneticProperties;
use phonetic_run::XlsxPhoneticRun;
use quick_xml::events::Event;
use rich_text_run::XlsxRichTextRun;
use text::load_text;

use crate::{common_types::Text, excel::XmlReader};

/// Example:
/// ```
/// // shared string
/// <si>
///     <t>課きく　毛こ</t>
///     <rPh sb="0" eb="1">
///         <t>カ</t>
///     </rPh>
///     <rPh sb="4" eb="5">
///        <t>ケ</t>
///     </rPh>
///     <phoneticPr fontId="1"/>
/// </si>
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
///
/// // inline string
/// <c r="A1" t="inlineStr">
///     <is><t>This is inline string example</t></is>
/// </c>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxStringItem {
    // Child Elements
    // phoneticPr (Phonetic Properties)	§18.4.3
    pub phonetic_properties: Option<XlsxPhoneticProperties>,
    // r (Rich Text Run)	§18.4.4
    pub rich_text_run: Option<Vec<XlsxRichTextRun>>,
    // rPh (Phonetic Run)	§18.4.6
    pub phonetic_run: Option<Vec<XlsxPhoneticRun>>,
    // t (Text)
    pub text: Option<Text>,
}

impl XlsxStringItem {
    pub(crate) fn load(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Self> {
        let mut item = Self {
            phonetic_properties: None,
            rich_text_run: None,
            phonetic_run: None,
            text: None,
        };
        let mut rich_text_runs: Vec<XlsxRichTextRun> = vec![];
        let mut phonetic_runs: Vec<XlsxPhoneticRun> = vec![];

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"phoneticPr" => {
                    item.phonetic_properties = Some(XlsxPhoneticProperties::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"r" => {
                    rich_text_runs.push(XlsxRichTextRun::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPh" => {
                    phonetic_runs.push(XlsxPhoneticRun::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" => {
                    item.text = Some(load_text(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!(
                    "unexpected end of file at `{}`.",
                    String::from_utf8(tag.to_vec()).unwrap_or("(unknown)".to_owned())
                ),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        item.phonetic_run = Some(phonetic_runs);
        item.rich_text_run = Some(rich_text_runs);

        return Ok(item);
    }
}
