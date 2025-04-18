use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::common_types::Dimension;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetdimension?view=openxml-3.0.1
///
/// This element specifies the used range of the worksheet.
/// It specifies the row and column bounds of used cells in the worksheet.
/// This is optional and is not required.
/// Used cells include cells with formulas, text content, and cell formatting.
/// When an entire column is formatted, only the first cell in that column is considered used.
///
/// Example:
/// ```
///
/// <dimension ref="A1:C2"/>
/// ```
/// dimension (Worksheet Dimensions)
pub type XlsxSheetDimension = Dimension;

pub(crate) fn load_sheet_dimension(e: &BytesStart) -> anyhow::Result<Option<XlsxSheetDimension>> {
    let attributes = e.attributes();

    for a in attributes {
        match a {
            Ok(a) => match a.key.local_name().as_ref() {
                b"ref" => {
                    let value = a.value.as_ref();
                    return Ok(Dimension::from_a1(value));
                }
                _ => {}
            },
            Err(error) => {
                bail!(error.to_string())
            }
        }
    }
    Ok(None)
}
