use anyhow::bail;

#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

use crate::{
    packaging::relationship::{zip_path_for_id, XlsxRelationships},
    raw::spreadsheet::workbook::sheet::XlsxSheet,
};

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct SheetBasicInfo {
    /// id (Relationship Id)
    pub r_id: String,

    /// name (Sheet Name)
    pub name: String,

    /// sheetId (Sheet Tab Id)
    pub sheet_id: u64,

    /// state (Visible State)
    pub visible_state: SheetVisibleState,

    /// type
    pub r#type: SheetType,

    /// xml path
    #[cfg_attr(feature = "serde", serde(skip_serializing, skip_deserializing))]
    pub(crate) path: String,
}

impl SheetBasicInfo {
    pub(crate) fn from_raw(
        sheet: XlsxSheet,
        relationships: &XlsxRelationships,
    ) -> anyhow::Result<Self> {
        let (Some(id), Some(name), Some(sheet_id)) = (sheet.id, sheet.name, sheet.sheet_id) else {
            bail!("neccessary properties for sheet are not present.")
        };
        let Some(path) = zip_path_for_id(relationships, &id) else {
            bail!("Cannot find the xml file for the sheet.")
        };

        let sheet_type = match path.split('/').nth(1) {
            Some("worksheets") => SheetType::WorkSheet,
            Some("chartsheets") => SheetType::ChartSheet,
            Some("dialogsheets") => SheetType::DialogSheet,
            Some(t) => bail!("Unsupported sheet type: {}", t),
            None => bail!("sheet type not availalbe."),
        };

        let visibility = match sheet.visible_state.unwrap_or("visible".to_owned()).as_ref() {
            "visible" => SheetVisibleState::Visible,
            "hidden" => SheetVisibleState::Hidden,
            "veryHidden" => SheetVisibleState::VeryHidden,
            state => bail!("unknown sheet visible state: {}", state),
        };
        return Ok(Self {
            r_id: id,
            name,
            sheet_id,
            visible_state: visibility,
            r#type: sheet_type,
            path,
        });
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub enum SheetType {
    /// WorkSheet
    WorkSheet,
    /// DialogSheet
    DialogSheet,
    /// ChartSheet
    ChartSheet,
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetstatevalues?view=openxml-3.0.1
#[derive(Debug, Clone, Copy, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub enum SheetVisibleState {
    /// Visible
    Visible,
    /// Hidden
    Hidden,
    /// The sheet is hidden and cannot be displayed using the user interface.
    VeryHidden,
}
