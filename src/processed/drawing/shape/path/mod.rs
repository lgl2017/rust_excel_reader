pub mod arc_to;
pub mod cubic_bezier_curve_to;
pub mod line_to;
pub mod move_to;
pub mod path_fill_mode;
pub mod path_shade_values;
pub mod path_type;
pub mod quad_bezier_curve_to;

use path_fill_mode::PathFillModeValues;
use path_type::PathTypeValues;
#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{
    shape::{path::XlsxPath, shape_guide::XlsxShapeGuide},
    st_types::emu_to_pt,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.path?view=openxml-3.0.1
///
/// This element specifies a creation path consisting of a series of moves, lines and curves that when combined forms a geometric shape
///
/// Example
/// ```
///   <a:pathLst>
///     <a:path w="3810000" h="3581400" fill="none" extrusionOk="0">
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
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Path {
    paths: Vec<PathTypeValues>,

    /// Specifies how the corresponding path should be filled.
    ///
    /// If this attribute is omitted, a value of "norm" is assumed.
    ///
    /// * Darken: Darken Path Fill.
    /// * DarkenLess: Darken Path Fill Less.
    /// * Lighten: Lighten Path Fill.
    /// * LightenLess: Lighten Path Fill Less.
    /// * None: No Path Fill.
    /// * Norm: Normal Path Fill.
    pub fill: PathFillModeValues,

    /// Specifies the width, or maximum x coordinate in points that should be used for within the path coordinate system.
    ///
    /// This value determines the horizontal placement of all points within the corresponding path as they are all calculated using this width attribute as the max x coordinate.
    ///
    /// (Not the line width)
    pub width: f64,

    /// Specifies the height, or maximum y coordinate in points that should be used for within the path coordinate system.
    ///
    /// This value determines the vertical placement of all points within the corresponding path as they are all calculated using this height attribute as the max y coordinate.
    pub height: f64,

    /// Specifies if the corresponding path should have a path stroke shown.
    ///
    /// This is a boolean value that affect the outline of the path.
    ///
    /// default to true.
    pub stroke: bool,

    /// Specifies that the use of 3D extrusions are possible on this path.
    /// This allows the generating application to know whether 3D extrusion can be applied in any form.
    ///
    /// If this attribute is omitted, then a value of 0, or false is assumed.
    pub extrusion_allowed: bool,
}

impl Path {
    pub(crate) fn from_raw(raw: XlsxPath, guide_list: Option<Vec<XlsxShapeGuide>>) -> Self {
        let paths: Vec<PathTypeValues> = raw
            .paths
            .unwrap_or(vec![])
            .into_iter()
            .map(|raw| PathTypeValues::from_raw(raw, guide_list.clone()))
            .collect();

        return Self {
            paths,
            fill: PathFillModeValues::from_string(raw.fill),
            width: emu_to_pt(raw.width.unwrap_or(0) as i64),
            height: emu_to_pt(raw.height.unwrap_or(0) as i64),
            stroke: raw.stroke.unwrap_or(true),
            extrusion_allowed: raw.extrusion_allowed.unwrap_or(false),
        };
    }
}
