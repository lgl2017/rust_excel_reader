use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_bool;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.drawing.applicationnonvisualdrawingproperties?view=openxml-3.0.1
///
/// A complex type that specifies SpreadsheetML Drawing-specific non-visual properties of a content part.
///
/// xdr14:nvPr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxApplicationNonVisualDrawingProperties {
    // attributes
    /// the name of the custom function associated with the content part.
    pub r#macro: Option<String>,

    /// specifies whether the content part is published with the worksheet when sent to the server.
    ///
    /// default to false
    pub f_published: Option<bool>,
}

impl XlsxApplicationNonVisualDrawingProperties {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut properties = Self {
            r#macro: None,
            f_published: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"macro" => {
                            properties.r#macro = Some(string_value);
                        }
                        b"fPublished" => {
                            properties.f_published = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        return Ok(properties);
    }
}
