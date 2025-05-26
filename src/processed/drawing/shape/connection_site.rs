#[cfg(feature = "serde")]
use serde::Serialize;

use crate::processed::drawing::common_types::adjust_angle::AdjustAngle;
use crate::processed::drawing::common_types::position::Position;
use crate::raw::drawing::shape::connection_site::XlsxConnectionSite;
use crate::raw::drawing::shape::shape_guide::XlsxShapeGuide;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionsite?view=openxml-3.0.1
///
/// This element specifies the existence of a connection site on a custom shape.
///
/// A connection site allows a cxnSp ([ConnectionShape](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionshape?view=openxml-3.0.1)) to be attached to this shape. This connection is maintained when the shape is repositioned within the document.
///
/// Example
/// ```
/// <a:cxnLst>
///     <a:cxn ang="0">
///         <a:pos x="0" y="679622"/>
///     </a:cxn>
///     <a:cxn ang="0">
///         <a:pos x="1705233" y="679622"/>
///     </a:cxn>
/// </a:cxnLst>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ConnectionSite {
    /// Position of the connection site
    pub position: Position,

    /// Specifies the incoming connector angle.
    ///
    /// This angle is the angle around the connection site that an incoming connector tries to be routed to.
    /// This allows connectors to know where the shape is in relation to the connection site and route connectors so as to avoid any overlap with the shape.
    pub angle: AdjustAngle,
}

impl ConnectionSite {
    pub(crate) fn from_raw(
        raw: XlsxConnectionSite,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        return Self {
            position: Position::from_position(raw.clone().position, guide_list.clone()),
            angle: AdjustAngle::from_raw(raw.angle, guide_list.clone()),
        };
    }
}
