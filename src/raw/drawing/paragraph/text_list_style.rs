use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use crate::raw::drawing::paragraph::paragraph_properties::{
    load_default_paragraph_properties, load_level1_paragraph_properties,
    load_level2_paragraph_properties, load_level3_paragraph_properties,
    load_level4_paragraph_properties, load_level5_paragraph_properties,
    load_level6_paragraph_properties, load_level7_paragraph_properties,
    load_level8_paragraph_properties, load_level9_paragraph_properties, DefaultParagraphProperties,
    Level1ParagraphProperties, Level2ParagraphProperties, Level3ParagraphProperties,
    Level4ParagraphProperties, Level5ParagraphProperties, Level6ParagraphProperties,
    Level7ParagraphProperties, Level8ParagraphProperties, Level9ParagraphProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.liststyle?view=openxml-3.0.1
///
/// Example
/// ```
/// <a:lstStyle>
///     <a:defPPr algn="ctr">
///         <a:defRPr sz="1600">
///             <a:solidFill>
///                 <a:schemeClr val="bg1" />
///             </a:solidFill>
///          </a:defRPr>
///     </a:defPPr>
/// </a:lstStyle>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct ListStyle {
    // child: extLst (Extension List) Not supported

    // Child Elements	Subclause
    // defPPr (Default Paragraph Style)	§21.1.2.2.2
    pub default_paragraph_style: Option<Box<DefaultParagraphProperties>>,

    // lvl1pPr (List Level 1 Text Style)	§21.1.2.4.13
    pub level1_paragraph_style: Option<Box<Level1ParagraphProperties>>,

    // lvl2pPr (List Level 2 Text Style)	§21.1.2.4.14
    pub level2_paragraph_style: Option<Box<Level2ParagraphProperties>>,

    // lvl3pPr (List Level 3 Text Style)	§21.1.2.4.15
    pub level3_paragraph_style: Option<Box<Level3ParagraphProperties>>,

    // lvl4pPr (List Level 4 Text Style)	§21.1.2.4.16
    pub level4_paragraph_style: Option<Box<Level4ParagraphProperties>>,

    // lvl5pPr (List Level 5 Text Style)	§21.1.2.4.17
    pub level5_paragraph_style: Option<Box<Level5ParagraphProperties>>,

    // lvl6pPr (List Level 6 Text Style)	§21.1.2.4.18
    pub level6_paragraph_style: Option<Box<Level6ParagraphProperties>>,

    // lvl7pPr (List Level 7 Text Style)	§21.1.2.4.19
    pub level7_paragraph_style: Option<Box<Level7ParagraphProperties>>,

    // lvl8pPr (List Level 8 Text Style)	§21.1.2.4.20
    pub level8_paragraph_style: Option<Box<Level8ParagraphProperties>>,

    // lvl9pPr (List Level 9 Text Style)
    pub level9_paragraph_style: Option<Box<Level9ParagraphProperties>>,
}

impl ListStyle {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut styles = Self {
            default_paragraph_style: None,
            level1_paragraph_style: None,
            level2_paragraph_style: None,
            level3_paragraph_style: None,
            level4_paragraph_style: None,
            level5_paragraph_style: None,
            level6_paragraph_style: None,
            level7_paragraph_style: None,
            level8_paragraph_style: None,
            level9_paragraph_style: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"defPPr" => {
                    styles.default_paragraph_style =
                        Some(Box::new(load_default_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl1pPr" => {
                    styles.level1_paragraph_style =
                        Some(Box::new(load_level1_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl2pPr" => {
                    styles.level2_paragraph_style =
                        Some(Box::new(load_level2_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl3pPr" => {
                    styles.level3_paragraph_style =
                        Some(Box::new(load_level3_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl4pPr" => {
                    styles.level4_paragraph_style =
                        Some(Box::new(load_level4_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl5pPr" => {
                    styles.level5_paragraph_style =
                        Some(Box::new(load_level5_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl6pPr" => {
                    styles.level6_paragraph_style =
                        Some(Box::new(load_level6_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl7pPr" => {
                    styles.level7_paragraph_style =
                        Some(Box::new(load_level7_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl8pPr" => {
                    styles.level8_paragraph_style =
                        Some(Box::new(load_level8_paragraph_properties(reader, e)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lvl9pPr" => {
                    styles.level9_paragraph_style =
                        Some(Box::new(load_level9_paragraph_properties(reader, e)?));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"lstStyle" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(styles)
    }
}
