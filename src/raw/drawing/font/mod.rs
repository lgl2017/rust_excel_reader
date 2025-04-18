use crate::excel::XmlReader;
use anyhow::bail;
use complex_sript_font::XlsxComplexScriptFont;
use east_asian_font::XlsxEastAsianFont;
use latin_font::XlsxLatinFont;
use quick_xml::events::Event;
use supplemental_font::XlsxSupplementalFont;

pub mod base_font;
pub mod complex_sript_font;
pub mod east_asian_font;
pub mod font_reference;
pub mod latin_font;
pub mod supplemental_font;
pub mod text_font_type;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.minorfont?view=openxml-3.0.1
///
/// Example:
/// ```
/// <minorFont>
///   <latin typeface="Calibri"/>
///   <ea typeface="Arial"/>
///   <cs typeface="Arial"/>
///   <font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
///   <font script="Hang" typeface="HY중고딕"/>
///   <font script="Hans" typeface="隶书"/>
///   <font script="Hant" typeface="微軟正黑體"/>
///   <font script="Arab" typeface="Traditional Arabic"/>
///   <font script="Hebr" typeface="Arial"/>
///   <font script="Thai" typeface="Cordia New"/>
///   <font script="Ethi" typeface="Nyala"/>
///   <font script="Beng" typeface="Vrinda"/>
///   <font script="Gujr" typeface="Shruti"/>
///   <font script="Khmr" typeface="DaunPenh"/>
///   <font script="Knda" typeface="Tunga"/>
/// </minorFont>
/// ```
pub type XlsxMinorFont = XlsxFontBase;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.majorfont?view=openxml-3.0.1
///
/// Example:
/// ```
/// <majorFont>
///   <latin typeface="Calibri"/>
///   <ea typeface="Arial"/>
///   <cs typeface="Arial"/>
///   <font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
///   <font script="Hang" typeface="HY중고딕"/>
///   <font script="Hans" typeface="隶书"/>
///   <font script="Hant" typeface="微軟正黑體"/>
///   <font script="Arab" typeface="Traditional Arabic"/>
///   <font script="Hebr" typeface="Arial"/>
///   <font script="Thai" typeface="Cordia New"/>
///   <font script="Ethi" typeface="Nyala"/>
///   <font script="Beng" typeface="Vrinda"/>
///   <font script="Gujr" typeface="Shruti"/>
///   <font script="Khmr" typeface="DaunPenh"/>
///   <font script="Knda" typeface="Tunga"/>
/// </majorFont>
/// ```
pub type XlsxMajorFont = XlsxFontBase;

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFontBase {
    // child: extLst (Extension List)	Not supported

    /* Children */
    /// ComplexScriptFont: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.complexscriptfont?view=openxml-3.0.1
    // cs (Complex Script Font), attribute: typeface
    pub cs: Option<XlsxComplexScriptFont>,

    /// EastAsianFont: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.eastasianfont?view=openxml-3.0.1
    // ea (East Asian Font), attribute: typeface
    pub ea: Option<XlsxEastAsianFont>,

    // font (Font)	§20.1.4.1.16
    pub font: Option<Vec<XlsxSupplementalFont>>,

    /// LatinFont: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.latinfont?view=openxml-3.0.1
    // tag: latin (Latin Font), attribute: typeface
    pub latin: Option<XlsxLatinFont>,
}

impl XlsxFontBase {
    pub(crate) fn load_major(reader: &mut XmlReader) -> anyhow::Result<XlsxMajorFont> {
        return Self::load_helper(reader, b"majorFont");
    }

    pub(crate) fn load_minor(reader: &mut XmlReader) -> anyhow::Result<XlsxMajorFont> {
        return Self::load_helper(reader, b"minorFont");
    }

    fn load_helper(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Self> {
        let mut font: Self = Self {
            cs: None,
            ea: None,
            font: None,
            latin: None,
        };
        let mut supplemental_fonts: Vec<XlsxSupplementalFont> = vec![];

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cs" => {
                    font.cs = Some(XlsxComplexScriptFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ea" => {
                    font.ea = Some(XlsxEastAsianFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                    supplemental_fonts.push(XlsxSupplementalFont::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"latin" => {
                    font.latin = Some(XlsxLatinFont::load(e)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        font.font = Some(supplemental_fonts);

        Ok(font)
    }
}
