#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::text::shape_text_rectangle::ShapeTextRectangle,
    raw::drawing::shape::custom_geometry::XlsxCustomGeometry,
};

use super::{
    adjust_handle_type::AdjustHandleTypeValues, connection_site::ConnectionSite, path::Path,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customgeometry?view=openxml-3.0.1
///
/// This element specifies the existence of a custom geometric shape.
///
/// Example
/// ```
/// <a:custGeom>
///   <a:avLst/>
///   <a:gdLst/>
///   <a:ahLst/>
///   <a:cxnLst>
///     <a:cxn ang="0">
///         <a:pos x="0" y="679622"/>
///     </a:cxn>
///     <a:cxn ang="0">
///         <a:pos x="1705233" y="679622"/>
///     </a:cxn>
///   </a:cxnLst>
///   <a:rect l="0" t="0" r="0" b="0"/>
///   <a:pathLst>
///     <a:path w="2650602" h="1261641">
///       <a:moveTo>
///         <a:pt x="0" y="1261641"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="2650602" y="1261641"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="1226916" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct CustomGeometry {
    /// List of Shape Paths
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathlist?view=openxml-3.0.1
    ///
    /// This element specifies the entire path that is to make up a single geometric shape.
    pub paths: Vec<Path>,

    /// Shape Text Rectangle
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectangle?view=openxml-3.0.1
    ///
    /// This element specifies the rectangular bounding box for text within a `custGeom` shape.
    /// The default for this rectangle is the bounding box for the shape.
    /// This can be modified using this elements four attributes to inset or extend the text bounding box.
    ///
    /// Example:
    /// ```
    /// <a:rect l="0" t="0" r="0" b="0"/>
    /// ```
    pub text_rectangle: ShapeTextRectangle,

    /// List of Shape Adjust Handles
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.adjusthandlelist?view=openxml-3.0.1
    ///
    /// This element specifies the adjust handles that are applied to a custom geometry.
    ///
    /// These adjust handles specify points within the geometric shape that can be used to perform certain transform operations on the shape.
    ///
    /// Example
    /// ```
    /// <a:ahLst>
    ///     <a:ahPolar gdRefAng="" gdRefR="">
    ///        <a:pos x="2" y="2"/>
    ///     </a:ahPolar>
    ///     <a:ahXY gdRefAng="" gdRefR="">
    ///        <a:pos x="2" y="2"/>
    ///     </a:ahXY>
    /// </a:ahLst>
    /// ```
    pub adjust_handles: Vec<AdjustHandleTypeValues>,

    /// List of Shape Connection Sites
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionsitelist?view=openxml-3.0.1
    ///
    /// This element specifies all the connection sites that are used for this shape.
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
    pub connection_sites: Vec<ConnectionSite>,
}

// // gdLst (List of Shape Guides)	§20.1.9.12
// pub guide_list: Option<XlsxShapeGuideList>,

// avLst (List of Shape Adjust Values)	§20.1.9.5
// pub adjust_value_list: Option<XlsxAdjustValueList>,

impl CustomGeometry {
    pub(crate) fn from_raw(raw: XlsxCustomGeometry) -> Self {
        let adv_list = raw.adjust_value_list.unwrap_or(vec![]);
        let sp_guide_list = raw.shape_guide_list.unwrap_or(vec![]);
        let guide_list = Some([&adv_list[..], &sp_guide_list[..]].concat());

        let paths: Vec<Path> = raw
            .path_list
            .unwrap_or(vec![])
            .into_iter()
            .map(|raw_path| Path::from_raw(raw_path, guide_list.clone()))
            .collect();

        let adjust_handles: Vec<AdjustHandleTypeValues> = raw
            .adjust_handle_list
            .unwrap_or(vec![])
            .into_iter()
            .map(|raw_handle| AdjustHandleTypeValues::from_raw(raw_handle, guide_list.clone()))
            .collect();

        let connection_sites: Vec<ConnectionSite> = raw
            .connection_site_list
            .unwrap_or(vec![])
            .into_iter()
            .map(|raw_cnx| ConnectionSite::from_raw(raw_cnx, guide_list.clone()))
            .collect();

        return Self {
            paths,
            text_rectangle: ShapeTextRectangle::from_raw(raw.text_rectangle, guide_list.clone()),
            adjust_handles,
            connection_sites,
        };
    }
}
