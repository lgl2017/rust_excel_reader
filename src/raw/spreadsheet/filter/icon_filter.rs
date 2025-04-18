use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_unsignedint;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.iconfilter?view=openxml-3.0.1
///
/// This element specifies the icon set and particular icon within that set to filter by.
/// For any cells whose icon does not match the specified criteria, the corresponding rows shall be hidden from view when the filter is applied.
///
/// Example:
/// ```
/// <filterColumn colId="3">
///   <iconFilter iconSet="3Arrows" iconId="0"/>
/// </filterColumn>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxIconFilter {
    // Attributes
    /// iconId (Icon Id)
    ///
    /// Zero-based index of an icon in an icon set.
    /// The absence of this attribute means "no icon".
    pub icon_id: Option<u64>,

    /// iconSet (Icon Set)
    ///
    /// Specifies which icon set is used in the filter criteria.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.iconsetvalues?view=openxml-3.0.1
    pub icon_set: Option<String>,
}

impl XlsxIconFilter {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut filter = Self {
            icon_id: None,
            icon_set: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"iconId" => filter.icon_id = string_to_unsignedint(&string_value),
                        b"iconSet" => filter.icon_set = Some(string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(filter)
    }
}
