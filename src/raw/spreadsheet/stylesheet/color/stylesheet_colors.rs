use anyhow::bail;
use quick_xml::events::Event;

use crate::{common_types::HexColor, excel::XmlReader, helper::format_hex_string};

use super::{rgb_color::XlsxRgbColor, XlsxColor};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colors?view=openxml-3.0.1
///
/// When the color palette is modified, the indexedColors collection is written.
/// When a custom color has been selected, the mruColors collection is written.
// tag: colors
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxStyleSheetColors {
    // children
    /// indexedColors
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.indexedcolors?view=openxml-3.0.1
    ///
    /// Example
    /// ```
    ///  <indexedColors>
    ///     <rgbColor rgb="ff000000" />
    ///     <rgbColor rgb="ffffffff" />
    ///  </indexedColors>
    /// ```
    pub indexed_colors: Vec<XlsxRgbColor>,

    // tag: mruColors
    pub mru_colors: Vec<XlsxColor>,
}

impl XlsxStyleSheetColors {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut colors = Self {
            indexed_colors: vec![],
            mru_colors: vec![],
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"indexedColors" => {
                    colors.indexed_colors = load_indexed_colors(reader)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"mruColors" => {
                    colors.mru_colors = load_mru_colors(reader)?;
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"colors" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(colors)
    }
}

impl XlsxStyleSheetColors {
    pub(crate) fn get_indexed_color(&self, index: u64) -> Option<HexColor> {
        let Ok(index) = TryInto::<usize>::try_into(index) else {
            return None;
        };
        let indexed_colors = self.indexed_colors.clone();

        if index < indexed_colors.len() {
            let color = indexed_colors[index].clone();
            if let Some(hex) = color.rgb {
                if let Ok(new) = format_hex_string(&hex, Some(true)) {
                    return Some(new);
                };
            }
        }

        return None;
    }
}

fn load_mru_colors(reader: &mut XmlReader) -> anyhow::Result<Vec<XlsxColor>> {
    let mut colors: Vec<XlsxColor> = vec![];
    let mut buf = Vec::new();

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"color" => {
                let color = XlsxColor::load(e)?;
                colors.push(color);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"mruColors" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    return Ok(colors);
}

fn load_indexed_colors(reader: &mut XmlReader) -> anyhow::Result<Vec<XlsxRgbColor>> {
    let mut colors: Vec<XlsxRgbColor> = vec![];
    let mut buf = Vec::new();

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rgbColor" => {
                let color = XlsxRgbColor::load(e)?;
                colors.push(color);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"indexedColors" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    return Ok(colors);
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.indexedcolors?view=openxml-3.0.1
///
/// Array of default mappings from indexed color value to ARGB value.
///
/// Note that 0-7 are redundant of 8-15 to preserve backwards compatibility.
/// When using the default indexed color palette, the values are not written out, but instead are implied.
/// When the color palette, has been modified from default, then the entire color palette is written out.
pub(crate) fn get_default_indexed_color_mapping() -> Vec<HexColor> {
    let default_mapping: Vec<String> = vec![
        "000000FF".to_ascii_lowercase(),
        "FFFFFFFF".to_ascii_lowercase(),
        "FF0000FF".to_ascii_lowercase(),
        "00FF00FF".to_ascii_lowercase(),
        "0000FFFF".to_ascii_lowercase(),
        "FFFF00FF".to_ascii_lowercase(),
        "FF00FFFF".to_ascii_lowercase(),
        "00FFFFFF".to_ascii_lowercase(),
        "000000FF".to_ascii_lowercase(),
        "FFFFFFFF".to_ascii_lowercase(),
        "FF0000FF".to_ascii_lowercase(),
        "00FF00FF".to_ascii_lowercase(),
        "0000FFFF".to_ascii_lowercase(),
        "FFFF00FF".to_ascii_lowercase(),
        "FF00FFFF".to_ascii_lowercase(),
        "00FFFFFF".to_ascii_lowercase(),
        "800000FF".to_ascii_lowercase(),
        "008000FF".to_ascii_lowercase(),
        "000080FF".to_ascii_lowercase(),
        "808000FF".to_ascii_lowercase(),
        "800080FF".to_ascii_lowercase(),
        "008080FF".to_ascii_lowercase(),
        "C0C0C0FF".to_ascii_lowercase(),
        "808080FF".to_ascii_lowercase(),
        "9999FFFF".to_ascii_lowercase(),
        "993366FF".to_ascii_lowercase(),
        "FFFFCCFF".to_ascii_lowercase(),
        "CCFFFFFF".to_ascii_lowercase(),
        "660066FF".to_ascii_lowercase(),
        "FF8080FF".to_ascii_lowercase(),
        "0066CCFF".to_ascii_lowercase(),
        "CCCCFFFF".to_ascii_lowercase(),
        "000080FF".to_ascii_lowercase(),
        "FF00FFFF".to_ascii_lowercase(),
        "FFFF00FF".to_ascii_lowercase(),
        "00FFFFFF".to_ascii_lowercase(),
        "800080FF".to_ascii_lowercase(),
        "800000FF".to_ascii_lowercase(),
        "008080FF".to_ascii_lowercase(),
        "0000FFFF".to_ascii_lowercase(),
        "00CCFFFF".to_ascii_lowercase(),
        "CCFFFFFF".to_ascii_lowercase(),
        "CCFFCCFF".to_ascii_lowercase(),
        "FFFF99FF".to_ascii_lowercase(),
        "99CCFFFF".to_ascii_lowercase(),
        "FF99CCFF".to_ascii_lowercase(),
        "CC99FFFF".to_ascii_lowercase(),
        "FFCC99FF".to_ascii_lowercase(),
        "3366FFFF".to_ascii_lowercase(),
        "33CCCCFF".to_ascii_lowercase(),
        "99CC00FF".to_ascii_lowercase(),
        "FFCC00FF".to_ascii_lowercase(),
        "FF9900FF".to_ascii_lowercase(),
        "FF6600FF".to_ascii_lowercase(),
        "666699FF".to_ascii_lowercase(),
        "969696FF".to_ascii_lowercase(),
        "003366FF".to_ascii_lowercase(),
        "339966FF".to_ascii_lowercase(),
        "003300FF".to_ascii_lowercase(),
        "333300FF".to_ascii_lowercase(),
        "993300FF".to_ascii_lowercase(),
        "993366FF".to_ascii_lowercase(),
        "333399FF".to_ascii_lowercase(),
        "333333FF".to_ascii_lowercase(),
        "000000FF".to_ascii_lowercase(),
        "FFFFFFFF".to_ascii_lowercase(),
    ];
    return default_mapping;
}
