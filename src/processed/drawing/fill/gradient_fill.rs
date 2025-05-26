use crate::common_types::HexColor;

use crate::processed::drawing::shape::path::path_shade_values::PathShadeValues;
use crate::raw::drawing::fill::gradient_fill::{
    XlsxGradientFill, XlsxGradientStop, XlsxLinearGradientFill, XlsxPathGradientFill,
};
use crate::raw::drawing::scheme::color_scheme::XlsxColorScheme;
use crate::raw::drawing::st_types::{st_angle_to_degree, st_percentage_to_float};

#[cfg(feature = "serde")]
use serde::Serialize;

use super::fill_rectangle::FillRectangle;
use super::tile_flip::TileFlipValues;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:gradFill rotWithShape="1">
///     <a:gsLst>
///         <a:gs pos="0">
///             <a:schemeClr val="phClr">
///                 <a:satMod val="103000" />
///                 <a:lumMod val="102000" />
///                 <a:tint val="94000" />
///             </a:schemeClr>
///         </a:gs>
///         <a:gs pos="100000">
///             <a:schemeClr val="phClr">
///                 <a:lumMod val="99000" />
///                 <a:satMod val="120000" />
///                 <a:shade val="78000" />
///             </a:schemeClr>
///         </a:gs>
///     </a:gsLst>
///     <a:lin ang="5400000" scaled="0" />
///
/// </a:gradFill>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct GradientFill {
    /// List of gradient stops that specifies the gradient colors and their relative positions in the color band.
    pub gradient_stops: Vec<GradientStop>,

    /// gradient type
    ///
    /// * Linear
    /// * Path (Radial, Rectangular, Shape Path)
    pub gradient_type: GradientFillTypeValues,

    /// This element specifies a rectangular region of the shape to which the gradient is applied.
    pub tile_rect: FillRectangle,

    /// Specifies the direction(s) in which to flip the gradient while tiling.
    ///
    /// * Horizontal
    /// * HorizontalAndVertical
    /// * None
    /// * Vertical
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tileflipvalues?view=openxml-3.0.1
    pub flip: TileFlipValues,

    /// Specifies if a fill rotates along with a shape when the shape is rotated.
    pub rotate_with_shape: bool,
}

impl GradientFill {
    pub(crate) fn from_raw(
        raw: Option<XlsxGradientFill>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };

        let Some(fill_type) = GradientFillTypeValues::from_raw(raw.clone()) else {
            return None;
        };

        let gs: Vec<GradientStop> = raw
            .gs_lst
            .clone()
            .unwrap_or(vec![])
            .into_iter()
            .map(|s| GradientStop::from_raw(s, color_scheme.clone(), ref_color.clone()))
            .collect();

        return Some(Self {
            gradient_stops: gs,
            gradient_type: fill_type,
            tile_rect: FillRectangle::from_raw(raw.tile_rect.clone()),
            flip: TileFlipValues::from_string(raw.flip.clone()),
            rotate_with_shape: raw.rotate_with_shape.clone().unwrap_or(false),
        });
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientstop?view=openxml-3.0.1
///
/// This element defines a gradient stop.
/// A gradient stop consists of a position where the stop appears in the color band.
///
/// Example:
/// ```
/// <a:gs pos="100000">
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:gs>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct GradientStop {
    /// color
    pub color: HexColor,

    /// position
    ///
    /// Specifies where this gradient stop should appear in the color band.
    /// This position is specified in the range [0%, 100%], which corresponds to the beginning and the end of the color band respectively.
    pub position: f64,
}

impl GradientStop {
    pub(crate) fn from_raw(
        raw: XlsxGradientStop,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Self {
        let hex = if let Some(color) = raw.color.clone() {
            color.to_hex(color_scheme.clone(), ref_color.clone())
        } else {
            None
        };
        return Self {
            color: hex.unwrap_or("00000000".to_string()),
            position: st_percentage_to_float(raw.pos.clone().unwrap_or(0) as i64),
        };
    }
}

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum GradientFillTypeValues {
    Linear(LinearGradientFill),
    Path(PathGradientFill),
}

impl GradientFillTypeValues {
    pub(crate) fn from_raw(raw: XlsxGradientFill) -> Option<Self> {
        if let Some(lin) = raw.lin.clone() {
            return Some(Self::Linear(LinearGradientFill::from_raw(lin)));
        };

        if let Some(path) = raw.path.clone() {
            return Some(Self::Path(PathGradientFill::from_raw(path)));
        }

        return None;
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:lin ang="5400000" scaled="0" />
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct LinearGradientFill {
    // attributes
    /// Specifies the direction of color change for the gradient.
    pub angle: f64,

    /// Whether the gradient angle scales with the fill region.
    ///
    /// Mathematically, if this flag is `true`, then the gradient vector ( cos x , sin x ) is scaled by the width (w) and height (h) of the fill region, so that the vector becomes ( w cos x, h sin x ) (before normalization).
    /// Observe that now if the gradient angle is 45 degrees, the gradient vector is ( w, h ), which goes from top-left to bottom-right of the fill region.
    ///
    /// If this flag is `false`, the gradient angle is independent of the fill region and is not scaled using the manipulation described above.
    /// So a 45-degree gradient angle always give a gradient band whose line of constant color is parallel to the vector (1, -1).
    pub scale_with_fill: bool,
}

impl LinearGradientFill {
    pub(crate) fn from_raw(raw: XlsxLinearGradientFill) -> Self {
        return Self {
            angle: st_angle_to_degree(raw.ang.clone().unwrap_or(0) as i64),
            scale_with_fill: raw.scaled.unwrap_or(false),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathgradientfill?view=openxml-3.0.1
///
/// This element defines that a gradient fill follows a path vs. a linear line.
///
/// Example:
/// ```
/// <a:path path="circle">
///     <a:fillToRect l="50000" t="-80000" r="50000" b="180000" />
/// </a:path>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct PathGradientFill {
    /// defines the "focus" rectangle for the center shade, specified relative to the fill tile rectangle
    ///
    /// Each edge of the center shade rectangle is defined by a percentage offset from the corresponding edge of the tile rectangle.
    /// A positive percentage specifies an inset, while a negative percentage specifies an outset.
    pub fill_to_rect: FillRectangle,

    /// Specifies the direction of color change for the gradient.
    ///
    /// * Circle (Radial)
    /// * Rectangle (Rectangular)
    /// * Shape (Shape Path)
    pub path: PathShadeValues,
}

impl PathGradientFill {
    pub(crate) fn from_raw(raw: XlsxPathGradientFill) -> Self {
        return Self {
            fill_to_rect: FillRectangle::from_raw(raw.fill_to_rect.clone()),
            path: PathShadeValues::from_string(raw.path.clone()),
        };
    }
}
