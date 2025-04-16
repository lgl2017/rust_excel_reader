use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{
    common_types::{Coordinate, Dimension},
    helper::a1_dimension_to_row_col,
};

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
pub type SheetDimension = Dimension;

pub(crate) fn load_sheet_dimension(e: &BytesStart) -> anyhow::Result<Option<SheetDimension>> {
    let attributes = e.attributes();

    for a in attributes {
        match a {
            Ok(a) => match a.key.local_name().as_ref() {
                b"ref" => {
                    let value = a.value.as_ref();
                    let dimension = a1_dimension_to_row_col(value)?;
                    return Ok(Some(SheetDimension {
                        start: Coordinate::from_point(dimension.0),
                        end: Coordinate::from_point(dimension.1),
                    }));
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

// #[derive(Debug, Default, PartialEq, Eq, Hash, Ord, PartialOrd, Copy, Clone)]
// pub struct SheetDimension {
//     // Attributes
//     /// ref (Reference)	T
//     ///
//     /// The row and column bounds of all cells in this worksheet.
//     /// Corresponds to the range that would contain all c elements written under sheetData.
//     /// Does not support whole column or whole row reference notation.
//     ///
//     /// Start and end are converted from A1 to R1C1
//     pub start: CellCoordinate,
//     pub end: CellCoordinate,
// }

// impl SheetDimension {
//     pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Option<Self>> {
//         let attributes = e.attributes();

//         for a in attributes {
//             match a {
//                 Ok(a) => match a.key.local_name().as_ref() {
//                     b"ref" => {
//                         let value = a.value.as_ref();
//                         let dimension = a1_dimension_to_row_col(value)?;
//                         return Ok(Some(Self {
//                             start: CellCoordinate::from_point(dimension.0),
//                             end: CellCoordinate::from_point(dimension.1),
//                         }));
//                     }
//                     _ => {}
//                 },
//                 Err(error) => {
//                     bail!(error.to_string())
//                 }
//             }
//         }
//         Ok(None)
//     }
// }
