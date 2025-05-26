use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_unsignedint;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.autonumberedbullet?view=openxml-3.0.1
///
/// This element specifies that automatic numbered bullet points should be applied to a paragraph.
///
/// Example:
/// ```
/// <a:pPr …>
///     <a:buAutoNum type="arabicPeriod"/>
/// </a:pPr>
/// ```
// tag: buAutoNum
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAutoNumberedBullet {
    // Attributes
    /// Specifies the number that starts a given sequence of automatically numbered bullets.
    /// When the numbering is alphabetical, the number should map to the appropriate letter.
    /// For instance 1 maps to 'a', 2 to 'b' and so on. If the numbers are larger than 26, then multiple letters should be used. For instance 27 should be represented as 'aa' and similarly 53 should be 'aaa'.
    // startAt (Start Numbering At)
    pub start_at: Option<u64>,

    /// Specifies the numbering scheme that is to be used.
    /// This allows for the describing of formats other than strictly numbers.
    /// For instance, a set of bullets can be represented by a series of Roman numerals instead of the standard 1,2,3,etc. number set.
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textautonumberschemevalues?view=openxml-3.0.1
    // type (Bullet Autonumbering Type)
    pub r#type: Option<String>,
}

impl XlsxAutoNumberedBullet {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut bullet = Self {
            start_at: None,
            r#type: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"startAt" => {
                            bullet.start_at = string_to_unsignedint(&string_value);
                        }
                        b"type" => {
                            bullet.r#type = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(bullet)
    }
}
