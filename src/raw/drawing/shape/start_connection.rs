use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_unsignedint;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.startconnection?view=openxml-3.0.1
///
/// This element specifies the starting connection that should be made by the corresponding connector shape.
/// This connects the head of the connector to the first shape.
///
/// Example:
/// ```
/// <a:stCxn id="3" idx="1" />
/// ```
///
/// stCxn (Connection Start)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxStartConnection {
    // Attributes
    /// id (Identifier)
    ///
    /// an unsigned integer Specifies the id of the shape to make the final connection to.
    pub id: Option<u64>,

    /// idx (Index)
    ///
    /// Specifies the index into the connection site table of the final connection shape.
    /// That is there are many connection sites on a shape and it shall be specified which connection site the corresponding connector shape should connect to.
    pub index: Option<u64>,
}

impl XlsxStartConnection {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut start_connection = Self {
            id: None,
            index: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"id" => {
                            start_connection.id = string_to_unsignedint(&string_value);
                        }
                        b"idx" => {
                            start_connection.index = string_to_unsignedint(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        return Ok(start_connection);
    }
}
