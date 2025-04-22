use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_unsignedint},
};

/// CellStyles: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyles?view=openxml-3.0.1
///
/// Example
/// ```
/// <cellStyles count="4">
///     <cellStyle name="20% - Accent1" xfId="3" builtinId="30"/>
///     <cellStyle name="Heading 1" xfId="2" builtinId="16"/>
///     <cellStyle name="Normal" xfId="0" builtinId="0"/>
///     <cellStyle name="Title" xfId="1" builtinId="15"/>
/// </cellStyles>
/// ```
pub type XlsxCellStyles = Vec<XlsxCellStyle>;

pub(crate) fn load_cell_styles(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxCellStyles> {
    let mut buf: Vec<u8> = Vec::new();
    let mut styles: Vec<XlsxCellStyle> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellStyle" => {
                let style = XlsxCellStyle::load(e)?;
                styles.push(style);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellStyles" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(styles)
}

/// CellStyle: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyle?view=openxml-3.0.1
///
/// This element represents the name and related formatting records for a named cell style in this workbook.
///
/// Annex H contains a listing of cellStyles whose corresponding formatting records are implied rather than explicitly saved in the file.
/// In this case, a builtinId attribute is written on the cellStyle record, but no corresponding formatting records are written.
/// For all built-in cell styles, the builtinId determines the style, not the name. For all cell styles, Normal is applied by default.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCellStyle {
    // attributes
    /// The index of a built-in cell style
    ///
    /// xml tag: builtinId
    pub builtin_id: Option<u64>,

    /// True indicates that this built-in cell style has been customized.
    // xml tag: customBuiltin
    pub custom_builtin: Option<bool>,

    pub hidden: Option<bool>,

    /// Indicates that this formatting is for an outline style .
    // tag: iLevel
    pub i_level: Option<u64>,

    /// The name of the cell style
    pub name: Option<String>,

    /// CellStyle.FormatId: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyle.formatid?view=openxml-3.0.1#documentformat-openxml-spreadsheet-cellstyle-formatid
    /// Zero-based index referencing an xf record in the cellStyleXfs collection.
    /// This is used to determine the formatting defined for this named cell style.
    // tag: xfId
    pub xf_id: Option<u64>,
}

impl XlsxCellStyle {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut style = Self {
            builtin_id: None,
            custom_builtin: None,
            hidden: None,
            i_level: None,
            name: None,
            xf_id: None,
        };
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"builtinId" => style.builtin_id = string_to_unsignedint(&string_value),
                        b"customBuiltin" => style.custom_builtin = string_to_bool(&string_value),
                        b"hidden" => style.hidden = string_to_bool(&string_value),
                        b"iLevel" => style.i_level = string_to_unsignedint(&string_value),
                        b"name" => style.name = Some(string_value),
                        b"xfId" => style.xf_id = string_to_unsignedint(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        return Ok(style);
    }
}
