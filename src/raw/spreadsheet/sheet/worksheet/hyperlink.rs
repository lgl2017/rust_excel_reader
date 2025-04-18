use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{common_types::Coordinate, excel::XmlReader};

pub type XlsxHyperlinks = Vec<XlsxHyperlink>;

pub(crate) fn load_hyperlinks(reader: &mut XmlReader) -> anyhow::Result<XlsxHyperlinks> {
    let mut hyperlinks: XlsxHyperlinks = vec![];

    let mut buf = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hyperlink" => {
                hyperlinks.push(XlsxHyperlink::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"hyperlinks" => break,
            Ok(Event::Eof) => bail!("unexpected end of file at `hyperlinks`."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(hyperlinks)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-3.0.1
///
/// Example:
/// ```
/// <hyperlinks>
///   <hyperlink ref="A11" r:id="rId1" tooltip="Search Page"/>
/// </hyperlinks>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxHyperlink {
    // Attributes
    /// display (Display String)
    ///
    /// Display string, if different from string in string table.
    /// This is a property on the hyperlink object, but does not need to appear in the spreadsheet application UI.
    pub display_string: Option<String>,

    /// id (Relationship Id)
    ///
    /// Relationship Id in this sheet's relationships part, expressing the target location of the resource.
    pub r_id: Option<String>,

    /// location (Location)	Location within target.
    ///
    /// If target is a workbook (or this workbook) this shall refer to a sheet and cell or a defined name.
    pub location: Option<String>,

    /// ref (Reference)
    ///
    /// Cell location of hyperlink on worksheet.
    pub r#ref: Option<Coordinate>,

    /// tooltip (Tool Tip)
    ///
    /// This is additional text to help the user understand more about the hyperlink.
    /// This can be displayed as hover text when the mouse is over the link.
    pub tooltip: Option<String>,
}

impl XlsxHyperlink {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut link = Self {
            display_string: None,
            r_id: None,
            location: None,
            r#ref: None,
            tooltip: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"display" => {
                            link.display_string = Some(string_value);
                        }
                        b"id" => {
                            link.r_id = Some(string_value);
                        }
                        b"location" => {
                            link.location = Some(string_value);
                        }
                        b"ref" => {
                            link.r#ref = Coordinate::from_a1(&a.value);
                        }
                        b"tooltip" => {
                            link.tooltip = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(link)
    }
}
