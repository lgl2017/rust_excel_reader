#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{effect::soft_edge::XlsxSoftEdge, st_types::emu_to_pt};

/// softEdge (Soft Edge Effect)
///
/// This element specifies a soft edge effect.
/// The edges of the shape are blurred, while the fill is not affected.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.softedge?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct SoftEdge {
    /// Specifies the radius of blur to apply to the edges.
    pub radius: f64,
}

impl SoftEdge {
    pub(crate) fn from_raw(raw: Option<XlsxSoftEdge>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            radius: emu_to_pt(raw.clone().rad.unwrap_or(0) as i64),
        });
    }
}
