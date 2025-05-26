#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::scene::scene_3d_type::XlsxScene3DType;

use super::backdrop::BackDrop;
use super::camera::Camera;
use super::light_rig::LightRig;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.scene3dtype?view=openxml-3.0.1
///
/// This element defines optional scene-level 3D properties to apply to an object.
///
/// Example
/// ```
/// <a:scene3d>
///     <a:backdrop>
///         <anchor x="123" y="23" z="10000"/>
///         <norm dx="123" dy="23" dz="10000"/>
///         <up dx="123" dy="23" dz="10000"/>
///     </a:backdrop>
///     <a:camera prst="orthographicFront">
///         <a:rot lat="19902513" lon="17826689" rev="1362739"/>
///     </a:camera>
///     <a:lightRig rig="twoPt" dir="t">
///         <a:rot lat="0" lon="0" rev="6000000"/>
///     </a:lightRig>
/// </<a:scene3d>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct Scene3DProperties {
    /// backdrop (Backdrop Plane)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub backdrop: Option<BackDrop>,

    /// camera (Camera)	§20.1.5.5
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub camera: Option<Camera>,

    /// lightRig (Light Rig)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub light_rig: Option<LightRig>,
}

impl Scene3DProperties {
    pub(crate) fn from_raw(raw: Option<XlsxScene3DType>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            backdrop: BackDrop::from_raw(raw.clone().backdrop),
            camera: Camera::from_raw(raw.clone().camera),
            light_rig: LightRig::from_raw(raw.clone().light_rig),
        });
    }
}
