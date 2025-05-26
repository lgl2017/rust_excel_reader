#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::shape::{
    adjust_handle_list::XlsxAdjustHandleType, shape_guide::XlsxShapeGuide,
};

use super::{adjust_handle_polar::AdjustHandlePolar, adjust_handle_xy::AdjustHandleXY};

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum AdjustHandleTypeValues {
    Polar(AdjustHandlePolar),
    XY(AdjustHandleXY),
}

impl AdjustHandleTypeValues {
    pub(crate) fn from_raw(
        raw: XlsxAdjustHandleType,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        return match raw {
            XlsxAdjustHandleType::Polar(p) => {
                Self::Polar(AdjustHandlePolar::from_raw(p, guide_list.clone()))
            }
            XlsxAdjustHandleType::XY(xy) => {
                Self::XY(AdjustHandleXY::from_raw(xy, guide_list.clone()))
            }
        };
    }
}
