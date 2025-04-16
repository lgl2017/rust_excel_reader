use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_unsignedint;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.phoneticproperties?view=openxml-3.0.1
///
/// This element represents a collection of phonetic properties that affect the display of phonetic text for this String Item (si).
/// Phonetic text is used to give hints as to the pronunciation of an East Asian language, and the hints are displayed as text within the spreadsheet cells across the top portion of the cell.
/// Since the phonetic hints are text, every phonetic hint is expressed as a phonetic run (rPh), and these properties specify how to display that phonetic run.
///
/// Example
/// ```
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
/// ```
///
/// The above example shows a String Item that displays some Japanese text "課きく　毛こ."
/// It also displays some phonetic text across the top of the cell.
/// The phonetic text character, "カ" is displayed over the "課" character and the phonetic text "ケ" is displayed above the "毛" character, using the font record in the style sheet at index 1.
// tag: phoneticPr
#[derive(Debug, Clone, PartialEq)]
pub struct PhoneticProperties {
    // Attributes
    /// Specifies how the text for the phonetic run is aligned across the top of the cells, with respect to the main text in the body of the cell.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.phoneticalignmentvalues?view=openxml-3.0.1
    // alignment (Alignment)
    pub alignment: Option<String>,

    /// An integer that is a zero-based index into the font record in the style sheet.
    /// Represents the font to be used to display this phonetic run.
    ///
    /// If this index is out of bounds, then the default font of the Normal style should be used in its place.
    /// This default font should be at index 0.
    // fontId (Font Id)
    pub font_id: Option<u64>,

    /// An enumeration which specifies which East Asian character set should be used to display the phonetic run
    ///
    /// possible values: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_PhoneticType_topic_ID0EKVRFB.html
    /// - halfwidthKatakana
    /// - fullwidthKatakana
    /// - Hiragana
    /// - noConversion
    // type (Character Type)
    pub r#type: Option<String>,
}

impl PhoneticProperties {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut properties = Self {
            alignment: None,
            font_id: None,
            r#type: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"alignment" => {
                            properties.alignment = Some(string_value);
                        }
                        b"fontId" => {
                            properties.font_id = string_to_unsignedint(&string_value);
                        }
                        b"type" => {
                            properties.r#type = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(properties)
    }
}
