#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::shape::shape_guide::XlsxShapeGuide;
use crate::raw::drawing::st_types::{emu_to_pt, STAdjustCoordinate};

/// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_AdjCoordinate_topic_ID0E14KNB.html
///
/// - `Coordinate`: one dimensional position or length in points
/// - `Formula`: Formula defined in shape guide for calculating the coordinate
#[derive(Debug, PartialEq, PartialOrd, Clone)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum AdjustCoordinate {
    /// shape guide formula
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeguide?view=openxml-3.0.1
    //
    // mainly used when present in custGeom(CustomGeometry) where a gdLst (ShapeGuideList) is defined
    Formula(String),

    /// coordinate
    ///
    /// Represents a one dimensional position or length in points
    Coordinate(f64),
}

impl AdjustCoordinate {
    pub(crate) fn default() -> Self {
        Self::Coordinate(0.0)
    }

    pub(crate) fn from_raw(
        adjust_coornidate: Option<STAdjustCoordinate>,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        let Some(raw) = adjust_coornidate else {
            return Self::default();
        };

        let guide_list = guide_list.unwrap_or(vec![]);

        let adjust_coornidate = match raw {
            STAdjustCoordinate::GuideName(name) => {
                let guides: Vec<XlsxShapeGuide> = guide_list
                    .into_iter()
                    .filter(|gd| gd.name.clone().unwrap_or(String::new()) == name)
                    .collect();
                match guides.first() {
                    Some(gd) => Self::Formula(gd.formula.clone().unwrap_or(String::new())),
                    None => Self::Coordinate(0.0),
                }
            }
            STAdjustCoordinate::Coordinate(emu) => Self::Coordinate(emu_to_pt(emu)),
        };

        return adjust_coornidate;
    }
}
