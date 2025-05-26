#[cfg(feature = "serde")]
use serde::Serialize;

use crate::processed::drawing::common_types::adjust_angle::AdjustAngle;
use crate::processed::drawing::common_types::adjust_coordinate::AdjustCoordinate;
use crate::processed::drawing::common_types::position::Position;
use crate::raw::drawing::shape::adjust_handle_polar::XlsxAdjustHandlePolar;
use crate::raw::drawing::shape::shape_guide::XlsxShapeGuide;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.adjusthandlepolar?view=openxml-3.0.1
///
/// This element specifies a polar adjust handle for a custom shape.
///
/// Example
/// ```
/// <a:ahLst>
///     <a:ahPolar gdRefAng="" gdRefR="">
///        <a:pos x="2" y="2"/>
///     </a:ahPolar>
/// </a:ahLst>
/// ```
// tag: ahPolar
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]

pub struct AdjustHandlePolar {
    /// Position of the adjust handle
    pub position: Position,

    /// Specifies the name of the guide that is updated with the adjustment angle from this adjust handle.
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub angle_guide_ref: Option<String>,

    /// Specifies the name of the guide that is updated with the adjustment radius from this adjust handle.
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub radial_guid_ref: Option<String>,

    /// Specifies the maximum angle position that is allowed for this adjustment handle.
    ///
    /// If max_angle_adjustment and min_angle_adjustment are equal, this adjust handle cannot move angularly
    pub max_angle_adjustment: AdjustAngle,

    /// Specifies the minimum angle position that is allowed for this adjustment handle.
    ///
    /// If max_angle_adjustment and min_angle_adjustment are equal, this adjust handle cannot move angularly
    pub min_angle_adjustment: AdjustAngle,

    /// Specifies the maximum radial position that is allowed for this adjustment handle.
    ///
    /// If max_radial_adjustment and min_radial_adjustment are equal, this adjust handle cannot move radially
    pub max_radial_adjustment: AdjustCoordinate,

    /// Specifies the minimum radial position that is allowed for this adjustment handle.
    ///
    /// If max_radial_adjustment and min_radial_adjustment are equal, this adjust handle cannot move radially
    pub min_radial_adjustment: AdjustCoordinate,
}

impl AdjustHandlePolar {
    pub(crate) fn from_raw(
        raw: XlsxAdjustHandlePolar,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        return Self {
            position: Position::from_position(raw.clone().position, guide_list.clone()),
            angle_guide_ref: raw.clone().angle_guide_ref,
            radial_guid_ref: raw.clone().radial_guid_ref,
            max_angle_adjustment: AdjustAngle::from_raw(
                raw.max_angle_adjustment,
                guide_list.clone(),
            ),
            min_angle_adjustment: AdjustAngle::from_raw(
                raw.min_angle_adjustment,
                guide_list.clone(),
            ),
            max_radial_adjustment: AdjustCoordinate::from_raw(
                raw.max_radial_adjustment,
                guide_list.clone(),
            ),
            min_radial_adjustment: AdjustCoordinate::from_raw(
                raw.min_radial_adjustment,
                guide_list.clone(),
            ),
        };
    }
}
