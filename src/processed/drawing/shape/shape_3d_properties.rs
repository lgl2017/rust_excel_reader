#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    common_types::HexColor,
    raw::drawing::{
        scheme::color_scheme::XlsxColorScheme, shape::shape_3d_type::XlsxShape3DType,
        st_types::emu_to_pt,
    },
};

use super::{bevel::Bevel, preset_material::PresetMaterialTypeValues};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shape3dtype?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:sp3d extrusionH="165100" contourW="50800" prstMaterial="plastic">
///   <a:bevelT w="254000" h="254000"/>
///   <a:bevelB w="254000" h="254000"/>
///   <a:extrusionClr>
///     <a:srgbClr val="FF0000"/>
///   </a:extrusionClr>
///   <a:contourClr>
///     <a:schemeClr val="accent3"/>
///   </a:contourClr>
/// </a:sp3d>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Shape3DProperties {
    /// contourClr (Contour Color) 020b0fff
    pub contour_color: HexColor,

    /// contour width 0
    pub contour_width: f64,

    // extrusionClr (Extrusion Color) 00000000
    pub extrusion_color: HexColor,

    /// Extrusion Height 0
    pub extrusion_height: f64,

    /// Shape Depth (distance from gound)
    pub shape_depth: f64,

    /// bevelB (Bottom Bevel)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub bottom_bevel: Option<Bevel>,

    /// bevelT (Top Bevel)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub top_bevel: Option<Bevel>,

    /// Preset Material Type
    pub preset_material: PresetMaterialTypeValues,
}

impl Shape3DProperties {
    pub(crate) fn from_raw(
        raw: Option<XlsxShape3DType>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };
        let mut contour_color = "020b0fff".to_string();
        let mut extrusion_color = "00000000".to_string();

        if let Some(c) = raw.clone().contour_clr {
            if let Some(hex) = c.to_hex(color_scheme.clone(), ref_color.clone()) {
                contour_color = hex
            }
        };

        if let Some(c) = raw.clone().extrusion_clr {
            if let Some(hex) = c.to_hex(color_scheme.clone(), ref_color.clone()) {
                extrusion_color = hex
            }
        };

        return Some(Self {
            contour_color,
            contour_width: emu_to_pt(raw.clone().contour_w.unwrap_or(0) as i64),
            extrusion_color,
            extrusion_height: emu_to_pt(raw.clone().extrusion_h.unwrap_or(0) as i64),
            shape_depth: emu_to_pt(raw.clone().z.unwrap_or(0)),
            bottom_bevel: Bevel::from_raw(raw.clone().bevel_b),
            top_bevel: Bevel::from_raw(raw.clone().bevel_t),
            preset_material: PresetMaterialTypeValues::from_string(raw.clone().prst_material),
        });
    }
}
