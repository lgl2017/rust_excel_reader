use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_int, raw::drawing::st_types::STCoordinate};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tabstop?view=openxml-3.0.1
///
/// This element specifies a single tab stop to be used on a line of text when there are one or more tab characters present within the text.
///
/// Example:
/// ```
/// <a:tabLst>
///     <a:tab pos="2292350" algn="l"/>
///     <a:tab pos="2627313" algn="l"/>
///     <a:tab pos="2743200" algn="l"/>
///     <a:tab pos="2974975" algn="l"/>
/// </a:tabLst>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTabStop {
    // attributes
    /// Specifies the alignment that is to be applied to text using this tab stop.
    /// If this attribute is omitted then the application default for the generating application.
    ///
    /// Possible Values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.texttabalignmentvalues?view=openxml-3.0.1
    // algn
    pub alignment: Option<String>,

    /// Specifies the position of the tab stop relative to the left margin.
    /// If this attribute is omitted then the application default for tab stops is used.
    // pos
    pub position: Option<STCoordinate>,
}

impl XlsxTabStop {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut stop = Self {
            alignment: None,
            position: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"algn" => {
                            stop.alignment = Some(string_value);
                        }
                        b"pos" => {
                            stop.position = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(stop)
    }
}
