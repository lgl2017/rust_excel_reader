use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_bool;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.protection?view=openxml-3.0.1
///
/// Contains protection properties associated with the cell.
/// Each cell has protection properties that can be set.
/// The cell protection properties do not take effect unless the sheet has been protected.
#[derive(Debug, Clone, PartialEq)]
pub struct Protection {
    // attributes
    /// A boolean value indicating if the cell is hidden.
    pub hidden: Option<bool>,

    /// A boolean value indicating if the cell is locked.
    pub locked: Option<bool>,
}

impl Protection {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut protection = Self {
            hidden: None,
            locked: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"hidden" => protection.hidden = string_to_bool(&string_value),
                        b"locked" => protection.locked = string_to_bool(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        return Ok(protection);
    }
}
