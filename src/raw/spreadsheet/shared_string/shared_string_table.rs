use std::io::{Read, Seek};

use anyhow::bail;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::{excel::xml_reader, helper::string_to_unsignedint};

use super::shared_string_item::{load_shared_string_item, SharedStringItem};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sharedstringtable?view=openxml-3.0.1
///
/// Root element of Shared String
///
/// Example:
/// ```
/// <sst uniqueCount="19" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
///     <si>
///         <t>Basic</t>
///     </si>
///     <si>
///         <r>
///             <rPr>
///                 <b val="1" />
///                 <sz val="10" />
///                 <color indexed="8" />
///                 <rFont val="Helvetica Neue" />
///             </rPr>
///             <t>u</t>
///         </r>
///         <r>
///             <rPr>
///                 <i val="1" />
///                 <sz val="10" />
///                 <color indexed="8" />
///                 <rFont val="Helvetica Neue" />
///             </rPr>
///             <t>italic</t>
///         </r>
///     </si>
/// </sst>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct SharedStringTable {
    // Child Elements
    // extLst (Future Feature Data Storage Area)	§18.2.10
    // si (String Item)
    pub string_item: Option<Vec<SharedStringItem>>,

    // Attributes
    /// An integer representing the total count of strings in the workbook.
    /// This count does not include any numbers, it counts only the total of text strings in the workbook.
    // count (String Count)
    pub count: Option<u64>,

    /// An integer representing the total count of unique strings in the Shared String Table.
    /// A string is unique even if it is a copy of another string, but has different formatting applied at the character level.
    // uniqueCount (Unique String Count)
    pub unique_count: Option<u64>,
}

impl SharedStringTable {
    pub(crate) fn load(zip: &mut ZipArchive<impl Read + Seek>) -> anyhow::Result<Self> {
        let path = "xl/sharedStrings.xml";

        let mut shared_string = Self {
            string_item: None,
            count: None,
            unique_count: None,
        };

        let Some(mut reader) = xml_reader(zip, path) else {
            return Ok(shared_string);
        };

        let mut items: Vec<SharedStringItem> = vec![];

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sst" => {
                    let attributes = e.attributes();
                    for a in attributes {
                        match a {
                            Ok(a) => {
                                let string_value = String::from_utf8(a.value.to_vec())?;
                                match a.key.local_name().as_ref() {
                                    b"count" => {
                                        shared_string.count = string_to_unsignedint(&string_value);
                                    }
                                    b"uniqueCount" => {
                                        shared_string.unique_count =
                                            string_to_unsignedint(&string_value);
                                    }
                                    _ => {}
                                }
                            }
                            Err(error) => {
                                bail!(error.to_string())
                            }
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"si" => {
                    items.push(load_shared_string_item(&mut reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sst" => break,
                Ok(Event::Eof) => break,
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        shared_string.string_item = Some(items);

        return Ok(shared_string);
    }
}
