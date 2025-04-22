use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_int},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.definednames?view=openxml-3.0.1
///
/// This element defines the collection of defined names for this workbook.
///
/// Example
/// ```
/// <definedNames>
///   <definedName name="NamedFormula" comment="Comment text for defined name.">SUM(Sheet3!$B$2:$B$9)</definedName>
///   <definedName name="NamedRange">Sheet3!$A$1:$C$12</definedName>
///   <definedName name="NamedRangeFromExternalReference" localSheetId="2" hidden="1">Sheet5!$A$1:$T$47</definedName>
/// </definedNames>
/// ```
/// definedNames (Defined Names)
pub type XlsxDefinedNames = Vec<XlsxDefinedName>;

pub(crate) fn load_defined_names(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxDefinedNames> {
    let mut names: XlsxDefinedNames = vec![];

    let mut buf = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"definedName" => {
                names.push(XlsxDefinedName::load(reader, e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"definedNames" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(names)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.definedname?view=openxml-3.0.1
///
/// This element defines the defined names that are defined within this workbook.
/// Defined names are descriptive text that is used to represents a cell, range of cells, formula, or constant value.
/// Use easy-to-understand names, such as Products, to refer to hard to understand ranges, such as Sales!C20:C30.
///
/// Example
/// ```
/// <definedName name="NamedRange">Sheet3!$A$1:$C$12</definedName>
/// ```
/// definedName (Defined Name)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDefinedName {
    /// Text
    pub value: Option<String>,

    /// Attributes
    /// Comment
    /// Represents the following attribute in the schema: comment
    pub comment: Option<String>,

    /// Custom Menu Text
    /// Represents the following attribute in the schema: customMenu
    pub custom_menu: Option<String>,

    /// Description
    /// Represents the following attribute in the schema: description
    pub description: Option<String>,

    /// Function
    /// Represents the following attribute in the schema: function
    pub function: Option<bool>,

    /// Function Group Id
    /// Represents the following attribute in the schema: functionGroupId
    pub function_group_id: Option<i64>,

    /// Help
    /// Represents the following attribute in the schema: help
    pub help: Option<String>,

    /// Hidden Name
    /// Represents the following attribute in the schema: hidden
    pub hidden: Option<bool>,

    /// Local Name Sheet Id
    /// Represents the following attribute in the schema: localSheetId
    pub local_sheet_id: Option<i64>,

    /// Defined Name
    /// Represents the following attribute in the schema: name
    pub name: Option<String>,

    /// Publish To Server
    /// Represents the following attribute in the schema: publishToServer
    pub publish_to_server: Option<bool>,

    /// Shortcut Key
    /// Represents the following attribute in the schema: shortcutKey
    pub shortcut_key: Option<String>,

    /// Status Bar
    /// Represents the following attribute in the schema: statusBar
    pub status_bar: Option<String>,

    /// Procedure
    /// Represents the following attribute in the schema: vbProcedure
    pub vb_procedure: Option<bool>,

    /// Workbook Parameter (Server)
    /// Represents the following attribute in the schema: workbookParameter
    pub workbook_parameter: Option<bool>,

    /// External Function
    /// Represents the following attribute in the schema: xlm
    pub external_function: Option<bool>,
}

impl XlsxDefinedName {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut defined_name = Self {
            value: None,
            comment: None,
            custom_menu: None,
            description: None,
            function: None,
            function_group_id: None,
            help: None,
            hidden: None,
            local_sheet_id: None,
            name: None,
            publish_to_server: None,
            shortcut_key: None,
            status_bar: None,
            vb_procedure: None,
            workbook_parameter: None,
            external_function: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"comment" => {
                            defined_name.comment = Some(string_value);
                        }
                        b"customMenu" => {
                            defined_name.custom_menu = Some(string_value);
                        }
                        b"description" => {
                            defined_name.description = Some(string_value);
                        }
                        b"function" => {
                            defined_name.function = string_to_bool(&string_value);
                        }
                        b"functionGroupId" => {
                            defined_name.function_group_id = string_to_int(&string_value);
                        }
                        b"help" => {
                            defined_name.help = Some(string_value);
                        }
                        b"hidden" => {
                            defined_name.hidden = string_to_bool(&string_value);
                        }
                        b"localSheetId" => {
                            defined_name.local_sheet_id = string_to_int(&string_value);
                        }
                        b"name" => {
                            defined_name.name = Some(string_value);
                        }
                        b"publishToServer" => {
                            defined_name.publish_to_server = string_to_bool(&string_value);
                        }
                        b"shortcutKey" => {
                            defined_name.shortcut_key = Some(string_value);
                        }
                        b"statusBar" => {
                            defined_name.status_bar = Some(string_value);
                        }
                        b"vbProcedure" => {
                            defined_name.vb_procedure = string_to_bool(&string_value);
                        }
                        b"workbookParameter" => {
                            defined_name.workbook_parameter = string_to_bool(&string_value);
                        }
                        b"xlm" => {
                            defined_name.external_function = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf = Vec::new();
        let mut text = String::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Text(t)) => text.push_str(&t.unescape()?),
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"definedName" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        defined_name.value = Some(text);

        Ok(defined_name)
    }
}
