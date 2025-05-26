use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.drawing?view=openxml-3.0.1
///
/// The sheet contains drawing components built on the drawingML platform
///
/// Example:
/// ```
/// <drawing r:id="rId2" />
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDrawing {
    // Attributes
    /// id (Relationship id)
    ///
    /// Relationship Id referencing a part containing drawingML definitions for this worksheet.
    pub id: String,
}

impl XlsxDrawing {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"id" => return Ok(Self { id: string_value }),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        bail!("id not specified for the drawing.")
    }
}
