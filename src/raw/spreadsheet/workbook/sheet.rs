use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{excel::XmlReader, helper::string_to_unsignedint};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheets?view=openxml-3.0.1
///
/// This element represents the collection of sheets in the workbook.
///
/// Example:
/// ```
/// <sheets>
///   <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
///   <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
///   <sheet name="Sheet5" sheetId="3" r:id="rId3"/>
///   <sheet name="Chart1" sheetId="4" r:id="rId4"/>
/// </sheets>
/// ```
/// sheets (Sheets)
pub type XlsxSheets = Vec<XlsxSheet>;

pub(crate) fn load_sheets(reader: &mut XmlReader) -> anyhow::Result<XlsxSheets> {
    let mut sheets: XlsxSheets = vec![];

    let mut buf = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheet" => {
                sheets.push(XlsxSheet::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheets" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(sheets)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheet?view=openxml-3.0.1
///
/// This element defines a sheet in this workbook.
/// Sheet data is stored in a separate part.
///
/// Example
/// ```
/// <sheet name="Sheet 1" sheetId="3" r:id="rId1" />
/// ```
/// sheet (Sheet Information)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSheet {
    // Attributes
    /// id (Relationship Id)
    ///
    /// Namespace: .../officeDocument/2006/relationships
    ///
    /// Specifies the identifier of the sheet part where the definition for this sheet is stored.
    /// This attribute is required.
    pub id: Option<String>,

    /// name (Sheet Name)
    ///
    /// Specifies the name of the sheet. This name shall be unique.
    /// This attribute is required.
    pub name: Option<String>,

    /// sheetId (Sheet Tab Id)
    ///
    /// Specifies the internal identifier for the sheet.
    /// This identifier shall be unique.
    /// This attribute is required.
    pub sheet_id: Option<u64>,

    /// state (Visible State)
    ///
    /// Specifies the visible state of this sheet.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetstatevalues?view=openxml-3.0.1.
    /// The default value for this attribute is "visible."
    pub visible_state: Option<String>,
}

impl XlsxSheet {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut sheet = Self {
            id: None,
            name: None,
            sheet_id: None,
            visible_state: Some("visible".to_owned()),
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"id" => {
                            sheet.id = Some(string_value);
                        }
                        b"name" => {
                            sheet.name = Some(string_value);
                        }
                        b"sheetId" => {
                            sheet.sheet_id = string_to_unsignedint(&string_value);
                        }
                        b"state" => {
                            sheet.visible_state = Some(string_value);
                        }

                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        if sheet.id.is_none() || sheet.name.is_none() || sheet.sheet_id.is_none() {
            bail!("Sheet does not contain required attributes.")
        }

        Ok(sheet)
    }
}
