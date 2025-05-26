#[cfg(feature = "serde")]
use serde::Serialize;

use crate::processed::drawing::common_types::adjust_coordinate::AdjustCoordinate;
use crate::processed::drawing::common_types::position::Position;
use crate::raw::drawing::shape::adjust_handle_xy::XlsxAdjustHandleXY;
use crate::raw::drawing::shape::shape_guide::XlsxShapeGuide;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.adjusthandlexy?view=openxml-3.0.1
///
/// This element specifies an XY-based adjust handle for a custom shape.
///
/// Example
/// ```
/// <a:ahLst>
///     <a:ahXY gdRefAng="" gdRefR="">
///        <a:pos x="2" y="2"/>
///     </a:ahXY>
/// </a:ahLst>
/// ```
// tag: ahPolar
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct AdjustHandleXY {
    /// Position of the adjust handle
    pub position: Position,

    /// Specifies the name of the guide that is updated with the adjustment x position from this adjust handle.
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub horizontal_guide_ref: Option<String>,

    /// Specifies the name of the guide that is updated with the adjustment y position from this adjust handle.
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub vertical_guide_ref: Option<String>,

    /// Specifies the maximum horizontal position that is allowed for this adjustment handle.
    ///
    /// If max_horizaontal_adjustment and min_horizaontal_adjustment are equal, this adjust handle cannot move in the x direction
    pub max_horizaontal_adjustment: AdjustCoordinate,

    /// Specifies the minimum horizontal position that is allowed for this adjustment handle.
    ///
    /// If max_horizaontal_adjustment and min_horizaontal_adjustment are equal, this adjust handle cannot move in the x direction
    pub min_horizaontal_adjustment: AdjustCoordinate,

    /// Specifies the maximum vertical position that is allowed for this adjustment handle.
    ///
    /// If max_vertical_adjustment and min_vertical_adjustment are equal, this adjust handle cannot move in the y direction
    pub max_vertical_adjustment: AdjustCoordinate,

    /// Specifies the minimum vertical position that is allowed for this adjustment handle.
    ///
    /// If max_vertical_adjustment and min_vertical_adjustment are equal, this adjust handle cannot move in the y direction
    pub min_vertical_adjustment: AdjustCoordinate,
}

impl AdjustHandleXY {
    pub(crate) fn from_raw(
        raw: XlsxAdjustHandleXY,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        return Self {
            position: Position::from_position(raw.clone().position, guide_list.clone()),
            horizontal_guide_ref: raw.clone().horizontal_guide_ref,
            vertical_guide_ref: raw.clone().vertical_guide_ref,
            max_horizaontal_adjustment: AdjustCoordinate::from_raw(
                raw.max_horizaontal_adjustment,
                guide_list.clone(),
            ),
            min_horizaontal_adjustment: AdjustCoordinate::from_raw(
                raw.min_horizaontal_adjustment,
                guide_list.clone(),
            ),
            max_vertical_adjustment: AdjustCoordinate::from_raw(
                raw.max_vertical_adjustment,
                guide_list.clone(),
            ),
            min_vertical_adjustment: AdjustCoordinate::from_raw(
                raw.min_vertical_adjustment,
                guide_list.clone(),
            ),
        };
    }
}
