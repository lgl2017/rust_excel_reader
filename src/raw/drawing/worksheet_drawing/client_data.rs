use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_bool;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.clientdata?view=openxml-3.0.1
///
/// This element is used to set certain properties related to a drawing element on the client spreadsheet application.
///
/// clientData (Client Data)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxClientData {
    // Attributes
    /// fLocksWithSheet (Locks With Sheet Flag)
    ///
    /// This attribute indicates whether to disable selection on drawing elements when the sheet is protected.
    pub f_locks_with_sheet: Option<bool>,

    /// fPrintsWithSheet (Prints With Sheet Flag)
    ///
    /// This attribute indicates whether to print drawing elements when printing the sheet.
    pub f_prints_with_sheet: Option<bool>,
}

impl XlsxClientData {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut client_data = Self {
            f_locks_with_sheet: None,
            f_prints_with_sheet: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"fLocksWithSheet" => {
                            client_data.f_locks_with_sheet = string_to_bool(&string_value);
                        }
                        b"fPrintsWithSheet" => {
                            client_data.f_prints_with_sheet = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        return Ok(client_data);
    }
}
