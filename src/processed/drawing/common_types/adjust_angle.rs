#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::shape::shape_guide::XlsxShapeGuide;
use crate::raw::drawing::st_types::{st_angle_to_degree, STAdjustAngle};

/// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_AdjAngle_topic_ID0EZWKNB.html
///
/// - `Angle`: An angle in degree.
/// - `Formula`: Formula defined in shape guide for calculating the coordinate
#[derive(Debug, PartialEq, PartialOrd, Clone)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum AdjustAngle {
    /// shape guide formula
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeguide?view=openxml-3.0.1
    //
    // mainly used when present in custGeom(CustomGeometry) where a gdLst (ShapeGuideList) is defined
    Formula(String),

    /// Angle
    ///
    /// An angle in degree.
    ///
    /// Positive angles are clockwise (i.e., towards the positive y axis);
    /// negative angles are counter-clockwise (i.e., towards the negative y axis).
    Angle(f64),
}

impl AdjustAngle {
    pub(crate) fn default() -> Self {
        Self::Angle(0.0)
    }

    pub(crate) fn from_raw(
        adjust_coornidate: Option<STAdjustAngle>,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        let Some(raw) = adjust_coornidate else {
            return Self::default();
        };

        let guide_list = guide_list.unwrap_or(vec![]);

        let adjust_coornidate = match raw {
            STAdjustAngle::GuideName(name) => {
                let guides: Vec<XlsxShapeGuide> = guide_list
                    .into_iter()
                    .filter(|gd| gd.name.clone().unwrap_or(String::new()) == name)
                    .collect();
                match guides.first() {
                    Some(gd) => Self::Formula(gd.formula.clone().unwrap_or(String::new())),
                    None => Self::Angle(0.0),
                }
            }
            STAdjustAngle::Angle(angle) => Self::Angle(st_angle_to_degree(angle)),
        };

        return adjust_coornidate;
    }
}
