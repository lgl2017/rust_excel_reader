use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.headend?view=openxml-3.0.1
///
/// Example
/// ```
/// <headEnd len="lg" type="arrowhead" w="sm"/>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxHeadEnd {
    // attributes
    /// Specifies the line end length in relation to the line width.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineendlengthvalues?view=openxml-3.0.1
    /// - lg: Large
    /// - med: Medium
    /// - sm: Small
    // tag: len (Length of Head/End)
    pub len: Option<String>,

    /// Specifies the line end decoration, such as a triangle or arrowhead.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineendvalues?view=openxml-3.0.1
    // type (Line Head/End Type)
    pub r#type: Option<String>,

    /// Specifies the line end length in relation to the line width.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineendwidthvalues?view=openxml-3.0.1
    /// - lg: Large
    /// - med: Medium
    /// - sm: Small
    // w (Width of Head/End)
    pub w: Option<String>,
}

impl XlsxHeadEnd {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut stop = Self {
            len: None,
            r#type: None,
            w: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"len" => {
                            stop.len = Some(string_value);
                        }
                        b"type" => {
                            stop.r#type = Some(string_value);
                        }
                        b"w" => {
                            stop.w = Some(string_value);
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
