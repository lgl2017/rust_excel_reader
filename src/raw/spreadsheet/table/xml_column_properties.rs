use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_unsignedint},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.xmlcolumnproperties?view=openxml-3.0.1
///
/// An element defining the XML column properties for a column.
/// This is only used for tables created from XML mappings.
///
/// Example
/// ```
/// <tableColumn id="1" uniqueName="SomeElement" name="SomeElement">
///   <xmlColumnPr mapId="1" xpath="/xml/foo/element" xmlDataType="string"/>
/// </tableColumn>
/// ```

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxXmlColumnProperties {
    // extLst (Future Feature Data Storage Area) Not supported

    // attributes
    /// denormalized (Denormalized)
    pub denormalized: Option<bool>,

    /// mapId (MapId)
    pub map_id: Option<u64>,

    /// xmlDataType (XmlDataType)
    ///
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.xmldatavalues?view=openxml-3.0.1
    pub xml_data_type: Option<String>,

    /// xpath (XPath)
    pub xml_path: Option<String>,
}

impl XlsxXmlColumnProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut properties = Self {
            denormalized: None,
            map_id: None,
            xml_data_type: None,
            xml_path: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"denormalized" => properties.denormalized = string_to_bool(&string_value),
                        b"mapId" => properties.map_id = string_to_unsignedint(&string_value),
                        b"xmlDataType" => properties.xml_data_type = Some(string_value),
                        b"xpath" => properties.xml_path = Some(string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"xmlColumnPr" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `xmlColumnPr`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
