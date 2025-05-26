#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::shape::shape_properties::XlsxShapeProperties;

use super::{custom_geometry::CustomGeometry, preset_geometry_shape::PresetShapeTypeValues};

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum GeometryTypeValues {
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customgeometry?view=openxml-3.0.1
    ///
    /// This element specifies the existence of a custom geometric shape.
    /// This shape consists of a series of lines and curves described within a creation path.
    /// In addition to this there can also be adjust values, guides, adjust handles, connection sites and an inscribed rectangle specified for this custom geometric shape.
    ///
    ///
    /// Example:
    /// ```
    /// <a:custGeom>
    ///     <a:avLst />
    ///     <a:gdLst>
    ///         <a:gd name="connsiteX0" fmla="*/ 0 w 3810000" />
    ///         <a:gd name="connsiteY0" fmla="*/ 0 h 3581400" />
    ///         <a:gd name="connsiteX1" fmla="*/ 3810000 w 3810000" />
    ///         <a:gd name="connsiteY1" fmla="*/ 0 h 3581400" />
    ///     </a:gdLst>
    ///     <a:ahLst />
    ///     <a:cxnLst>
    ///         <a:cxn ang="0">
    ///             <a:pos x="connsiteX0" y="connsiteY0" />
    ///         </a:cxn>
    ///         <a:cxn ang="0">
    ///             <a:pos x="connsiteX1" y="connsiteY1" />
    ///         </a:cxn>
    ///     </a:cxnLst>
    ///     <a:rect l="l" t="t" r="r" b="b" />
    ///     <a:pathLst>
    ///         <a:path w="3810000" h="3581400" fill="none" extrusionOk="0">
    ///             <a:moveTo>
    ///                 <a:pt x="0" y="0" />
    ///             </a:moveTo>
    ///             <a:cubicBezTo>
    ///                 <a:pt x="1884255" y="-49533" />
    ///                 <a:pt x="2614916" y="-14809" />
    ///                 <a:pt x="3810000" y="0" />
    ///             </a:cubicBezTo>
    ///             <a:close />
    ///         </a:path>
    ///         <a:path w="3810000" h="3581400" stroke="0" extrusionOk="0">
    ///             <a:moveTo>
    ///                 <a:pt x="0" y="0" />
    ///             </a:moveTo>
    ///             <a:cubicBezTo>
    ///                 <a:pt x="1165673" y="118645" />
    ///                 <a:pt x="2493217" y="116012" />
    ///                 <a:pt x="3810000" y="0" />
    ///             </a:cubicBezTo>
    ///             <a:close />
    ///         </a:path>
    ///     </a:pathLst>
    /// </a:custGeom>
    /// ```
    CustomGeometry(CustomGeometry),

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetgeometry?view=openxml-3.0.1
    ///
    /// This element specifies when a preset geometric shape should be used instead of a custom geometric shape.
    PresetGeometry(PresetShapeTypeValues),
}

impl GeometryTypeValues {
    pub(crate) fn default() -> Self {
        Self::PresetGeometry(PresetShapeTypeValues::default())
    }

    pub(crate) fn from_shape_properties(raw: Option<XlsxShapeProperties>) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };

        if let Some(preset) = raw.clone().preset_gemoetry {
            return Self::PresetGeometry(PresetShapeTypeValues::from_string(preset.preset));
        };

        if let Some(custom) = raw.clone().custom_geometry {
            return Self::CustomGeometry(CustomGeometry::from_raw(custom));
        }

        return Self::default();
    }
}
