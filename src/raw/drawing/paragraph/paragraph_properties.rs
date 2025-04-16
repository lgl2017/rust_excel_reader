use std::collections::BTreeMap;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::raw::drawing::paragraph::{
    auto_numbered_bullet::AutoNumberedBullet, bullet_color::BulletColor,
    bullet_color_text::BulletColorText, bullet_font::BulletFont, bullet_font_text::BulletFontText,
    bullet_size_percentage::BulletSizePercentage, bullet_size_points::BulletSizePoints,
    bullet_size_text::BulletSizeText, character_bullet::CharacterBullet, no_bullet::NoBullet,
    picture_bullet::PictureBullet,
};

use super::{
    default_text_run_properties::DefaultTextRunProperties,
    line_spacing::LineSpacing,
    space_after::SpaceAfter,
    space_before::SpaceBefore,
    tab_stop_list::{load_tab_stop_list, TabStopList},
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
pub type DefaultParagraphProperties = ParagraphProperties;

pub(crate) fn load_default_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<DefaultParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"defPPr");
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
pub type Level1ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level1_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level1ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl1pPr");
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
pub type Level2ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level2_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level2ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl2pPr");
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
pub type Level3ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level3_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level3ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl3pPr");
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
pub type Level4ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level4_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level4ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl4pPr");
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
pub type Level5ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level5_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level5ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl5pPr");
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
pub type Level6ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level6_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level6ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl6pPr");
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
pub type Level7ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level7_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level7ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl7pPr");
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
pub type Level8ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level8_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level8ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl8pPr");
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
pub type Level9ParagraphProperties = ParagraphProperties;

pub(crate) fn load_level9_paragraph_properties(
    reader: &mut XmlReader,
    e: &BytesStart,
) -> anyhow::Result<Level9ParagraphProperties> {
    return ParagraphProperties::load(reader, e, b"lvl9pPr");
}

#[derive(Debug, Clone, PartialEq)]
pub struct ParagraphProperties {
    // child: extLst (Extension List) Not supported

    //  Child Elements
    // buAutoNum (Auto-Numbered Bullet)	§21.1.2.4.1
    pub auto_numbered_bullet: Option<AutoNumberedBullet>,

    // buBlip (Picture Bullet)	§21.1.2.4.2
    pub picture_bullet: Option<PictureBullet>,

    // buChar (Character Bullet)	§21.1.2.4.3
    pub character_bullet: Option<CharacterBullet>,

    // buClr (Color Specified)	§21.1.2.4.4
    pub bullet_color: Option<BulletColor>,

    // buClrTx (Follow Text)	§21.1.2.4.5
    pub bullet_color_text: Option<BulletColorText>,

    // buFont (Specified)	§21.1.2.4.6
    pub bullet_font: Option<BulletFont>,

    // buFontTx (Follow text)	§21.1.2.4.7
    pub bullet_font_text: Option<BulletFontText>,

    // buNone (No Bullet)	§21.1.2.4.8
    pub no_bullet: Option<NoBullet>,

    // buSzPct (Bullet Size Percentage)	§21.1.2.4.9
    pub bullet_size_percentage: Option<BulletSizePercentage>,

    // buSzPts (Bullet Size Points)	§21.1.2.4.10
    pub bullet_size_points: Option<BulletSizePoints>,

    // buSzTx (Bullet Size Follows Text)	§21.1.2.4.11
    pub bullet_size_text: Option<BulletSizeText>,

    // defRPr (Default Text Run Properties)	§21.1.2.3.2
    pub default_run_properties: Option<DefaultTextRunProperties>,

    // lnSpc (Line Spacing)	§21.1.2.2.5
    pub line_spacing: Option<LineSpacing>,

    // spcAft (Space After)	§21.1.2.2.9
    pub space_after: Option<SpaceAfter>,

    // spcBef (Space Before)	§21.1.2.2.10
    pub space_before: Option<SpaceBefore>,

    // tabLst (Tab List)
    pub tab_list: Option<TabStopList>,

    // attributes: undocumented
    pub extended_attributes: Option<BTreeMap<String, String>>,
}

impl ParagraphProperties {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart, tag: &[u8]) -> anyhow::Result<Self> {
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
            // undocumented attributes
            extended_attributes: None,
        };
        let mut attributes: BTreeMap<String, String> = BTreeMap::new();

        for a in e.attributes() {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    let key = String::from_utf8(a.key.local_name().as_ref().to_vec())?;
                    attributes.insert(key, string_value);
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        properties.extended_attributes = Some(attributes);

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buAutoNum" => {
                    properties.auto_numbered_bullet = Some(AutoNumberedBullet::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buBlip" => {
                    properties.picture_bullet = Some(PictureBullet::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buChar" => {
                    properties.character_bullet = Some(CharacterBullet::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buClr" => {
                    properties.bullet_color = BulletColor::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buClrTx" => {
                    properties.bullet_color_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buFont" => {
                    properties.bullet_font = Some(BulletFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buFontTx" => {
                    properties.bullet_font_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buNone" => {
                    properties.no_bullet = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buSzPct" => {
                    properties.bullet_size_percentage = Some(BulletSizePercentage::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buSzPts" => {
                    properties.bullet_size_points = Some(BulletSizePoints::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"buSzTx" => {
                    properties.bullet_size_text = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"defRPr" => {
                    properties.default_run_properties =
                        Some(DefaultTextRunProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lnSpc" => {
                    properties.line_spacing = LineSpacing::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spcAft" => {
                    properties.space_after = SpaceAfter::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spcBef" => {
                    properties.space_before = SpaceBefore::load(reader, e)?;
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
