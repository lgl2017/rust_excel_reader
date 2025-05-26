use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::excel::XmlReader;

use super::{
    auto_numbered_bullet::XlsxAutoNumberedBullet, bullet_color::XlsxBulletColor,
    bullet_color_text::XlsxBulletColorText, bullet_font::XlsxBulletFont,
    bullet_font_text::XlsxBulletFontText, bullet_size_percentage::XlsxBulletSizePercentage,
    bullet_size_points::XlsxBulletSizePoints, bullet_size_text::XlsxBulletSizeText,
    character_bullet::XlsxCharacterBullet, no_bullet::XlsxNoBullet,
    picture_bullet::XlsxPictureBullet,
};
use crate::helper::{string_to_bool, string_to_int, string_to_unsignedint};
use crate::raw::drawing::st_types::STCoordinate;

use super::{
    super::default_text_run_properties::XlsxDefaultTextRunProperties,
    line_spacing::XlsxLineSpacing,
    space_after::XlsxSpaceAfter,
    space_before::XlsxSpaceBefore,
    tab_stop_list::{load_tab_stop_list, XlsxTabStopList},
};

// There are a total of 9 level text property elements allowed, levels 0-8.

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.defaultparagraphproperties?view=openxml-3.0.1
///
/// This element specifies the paragraph properties that are to be applied when no other paragraph properties have been specified.
///
/// Example
/// ```
/// <a:defPPr>
///     <a:buNone/>
/// </a:defPPr>
/// ```
// tag: defPPr
pub type XlsxDefaultParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_default_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxDefaultParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"defPPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level1paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="0".
///
/// Example
/// ```
/// <a:lvl1pPr algn="r">
///     <a:buNone/>
/// </a:lvl1pPr>
/// ```
// tag: lvl1pPr
pub type XlsxLevel1ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level1_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel1ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl1pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level2paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="1".
///
/// Example
/// ```
/// <a:lvl2pPr algn="r">
///     <a:buNone/>
/// </a:lvl2pPr>
/// ```
// tag: lvl2pPr
pub type XlsxLevel2ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level2_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel2ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl2pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level3paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="2".
///
/// Example
/// ```
/// <a:lvl3pPr algn="r">
///     <a:buNone/>
/// </a:lvl3pPr>
/// ```
// tag: lvl3pPr
pub type XlsxLevel3ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level3_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel3ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl3pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level4paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="3".
///
/// Example
/// ```
/// <a:lvl4pPr algn="r">
///     <a:buNone/>
/// </a:lvl4pPr>
/// ```
// tag: lvl4pPr
pub type XlsxLevel4ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level4_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel4ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl4pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level5paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="4".
///
/// Example
/// ```
/// <a:lvl5pPr algn="r">
///     <a:buNone/>
/// </a:lvl5pPr>
/// ```
// tag: lvl5pPr
pub type XlsxLevel5ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level5_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel5ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl5pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level6paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="5".
///
/// Example
/// ```
/// <a:lvl6pPr algn="r">
///     <a:buNone/>
/// </a:lvl6pPr>
/// ```
// tag: lvl6pPr
pub type XlsxLevel6ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level6_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel6ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl6pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level7paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="6".
///
/// Example
/// ```
/// <a:lvl7pPr algn="r">
///     <a:buNone/>
/// </a:lvl7pPr>
/// ```
// tag: lvl7pPr
pub type XlsxLevel7ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level7_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel7ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl7pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level8paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="7".
///
/// Example
/// ```
/// <a:lvl8pPr algn="r">
///     <a:buNone/>
/// </a:lvl8pPr>
/// ```
// tag: lvl8pPr
pub type XlsxLevel8ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level8_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel8ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl8pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.level9paragraphproperties?view=openxml-3.0.1
///
/// This element specifies all paragraph level text properties for all elements that have the attribute lvl="8".
///
/// Example
/// ```
/// <a:lvl9pPr algn="r">
///     <a:buNone/>
/// </a:lvl9pPr>
/// ```
// tag: lvl9pPr
pub type XlsxLevel9ParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_level9_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxLevel9ParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"lvl9pPr");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraphproperties?view=openxml-3.0.1
///
/// pPr (Text Paragraph Properties)
pub type XlsxTextParagraphProperties = XlsxParagraphProperties;

pub(crate) fn load_text_paragraph_properties(
    reader: &mut XmlReader<impl Read>,

    e: &BytesStart,
) -> anyhow::Result<XlsxTextParagraphProperties> {
    return XlsxParagraphProperties::load(reader, e, b"pPr");
}

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxParagraphProperties {
    // child: extLst (Extension List) Not supported

    //  Child Elements
    // buAutoNum (Auto-Numbered Bullet)	§21.1.2.4.1
    pub auto_numbered_bullet: Option<XlsxAutoNumberedBullet>,

    // buBlip (Picture Bullet)	§21.1.2.4.2
    pub picture_bullet: Option<XlsxPictureBullet>,

    // buChar (Character Bullet)	§21.1.2.4.3
    pub character_bullet: Option<XlsxCharacterBullet>,

    // buClr (Color Specified)	§21.1.2.4.4
    pub bullet_color: Option<XlsxBulletColor>,

    // buClrTx (Follow Text)	§21.1.2.4.5
    pub bullet_color_text: Option<XlsxBulletColorText>,

    // buFont (Specified)	§21.1.2.4.6
    pub bullet_font: Option<XlsxBulletFont>,

    // buFontTx (Follow text)	§21.1.2.4.7
    pub bullet_font_text: Option<XlsxBulletFontText>,

    // buNone (No Bullet)	§21.1.2.4.8
    pub no_bullet: Option<XlsxNoBullet>,

    // buSzPct (Bullet Size Percentage)	§21.1.2.4.9
    pub bullet_size_percentage: Option<XlsxBulletSizePercentage>,

    // buSzPts (Bullet Size Points)	§21.1.2.4.10
    pub bullet_size_points: Option<XlsxBulletSizePoints>,

    // buSzTx (Bullet Size Follows Text)	§21.1.2.4.11
    pub bullet_size_text: Option<XlsxBulletSizeText>,

    // defRPr (Default Text Run Properties)	§21.1.2.3.2
    pub default_run_properties: Option<XlsxDefaultTextRunProperties>,

    // lnSpc (Line Spacing)	§21.1.2.2.5
    pub line_spacing: Option<XlsxLineSpacing>,

    // spcAft (Space After)	§21.1.2.2.9
    pub space_after: Option<XlsxSpaceAfter>,

    // spcBef (Space Before)	§21.1.2.2.10
    pub space_before: Option<XlsxSpaceBefore>,

    // tabLst (Tab List)
    pub tab_list: Option<XlsxTabStopList>,

    // attributes:
    // Extended from TextParagraphPropertiesType: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype?view=openxml-3.0.1
    /// Alignment
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textalignmenttypevalues?view=openxml-3.0.1
    pub text_alignment: Option<String>,

    /// DefaultTabSize
    pub default_tab_size: Option<STCoordinate>,

    /// EastAsianLineBreak
    pub east_asian_line_break: Option<bool>,

    /// FontAlignment
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textfontalignmentvalues?view=openxml-3.0.1
    pub font_alignment: Option<String>,

    /// Height
    pub hanging_punctuation: Option<bool>,

    /// Indent
    ///
    /// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextIndent_topic_ID0EM2QOB.html
    pub indent: Option<STCoordinate>,

    /// LatinLineBreak
    pub latin_line_break: Option<bool>,

    /// LeftMargin
    pub left_margin: Option<STCoordinate>,

    /// Level:
    ///
    /// This type specifies the indent level type. We support list level 0 to 8, and we use -1 and -2 for outline mode levels that should only exist in memory.
    ///
    /// - a minimum value of greater than or equal to 0.
    /// - a maximum value of less than or equal to 8.
    ///
    /// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextIndentLevelTy_topic_ID0EMGROB.html
    pub indent_level: Option<u64>,

    /// RightMargin
    pub right_margin: Option<STCoordinate>,

    /// RightToLeft
    pub right_to_left: Option<bool>,
}

impl XlsxParagraphProperties {
    pub(crate) fn load(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
        tag: &[u8],
    ) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut properties = Self {
            auto_numbered_bullet: None,
            picture_bullet: None,
            character_bullet: None,
            bullet_color: None,
            bullet_color_text: None,
            bullet_font: None,
            bullet_font_text: None,
            no_bullet: None,
            bullet_size_percentage: None,
            bullet_size_points: None,
            bullet_size_text: None,
            default_run_properties: None,
            line_spacing: None,
            space_after: None,
            space_before: None,
            tab_list: None,
            // attributes Extended from TextParagraphPropertiesType:
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype?view=openxml-3.0.1
            text_alignment: None,
            default_tab_size: None,
            east_asian_line_break: None,
            font_alignment: None,
            hanging_punctuation: None,
            indent: None,
            latin_line_break: None,
            left_margin: None,
            indent_level: None,
            right_margin: None,
            right_to_left: None,
        };

        for a in e.attributes() {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"algn" => {
                            properties.text_alignment = Some(string_value);
                        }
                        b"defTabSz" => {
                            properties.default_tab_size = string_to_int(&string_value);
                        }
                        b"eaLnBrk" => {
                            properties.east_asian_line_break = string_to_bool(&string_value);
                        }
                        b"fontAlgn" => {
                            properties.font_alignment = Some(string_value);
                        }
                        b"hangingPunct" => {
                            properties.hanging_punctuation = string_to_bool(&string_value);
                        }
                        b"indent" => {
                            properties.indent = string_to_int(&string_value);
                        }
                        b"latinLnBrk" => {
                            properties.latin_line_break = string_to_bool(&string_value);
                        }
                        b"marL" => {
                            properties.left_margin = string_to_int(&string_value);
                        }
                        b"lvl" => {
                            properties.indent_level = string_to_unsignedint(&string_value);
                        }
                        b"marR" => {
                            properties.right_margin = string_to_int(&string_value);
                        }
                        b"rtl" => {
                            properties.right_to_left = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buAutoNum" => {
                    properties.auto_numbered_bullet = Some(XlsxAutoNumberedBullet::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buBlip" => {
                    properties.picture_bullet = Some(XlsxPictureBullet::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buChar" => {
                    properties.character_bullet = Some(XlsxCharacterBullet::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buClr" => {
                    properties.bullet_color = XlsxBulletColor::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buClrTx" => {
                    properties.bullet_color_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buFont" => {
                    properties.bullet_font = Some(XlsxBulletFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buFontTx" => {
                    properties.bullet_font_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buNone" => {
                    properties.no_bullet = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buSzPct" => {
                    properties.bullet_size_percentage = Some(XlsxBulletSizePercentage::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buSzPts" => {
                    properties.bullet_size_points = Some(XlsxBulletSizePoints::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buSzTx" => {
                    properties.bullet_size_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"defRPr" => {
                    properties.default_run_properties =
                        Some(XlsxDefaultTextRunProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lnSpc" => {
                    properties.line_spacing = XlsxLineSpacing::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spcAft" => {
                    properties.space_after = XlsxSpaceAfter::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spcBef" => {
                    properties.space_before = XlsxSpaceBefore::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tabLst" => {
                    properties.tab_list = Some(load_tab_stop_list(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(properties)
    }
}
