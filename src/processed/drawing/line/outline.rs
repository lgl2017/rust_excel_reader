#[cfg(feature = "serde")]
use serde::Serialize;

use crate::common_types::HexColor;

use crate::processed::drawing::{common_types::pen_alignment::PenAlignmentValues, fill::Fill};

use crate::raw::drawing::{
    line::outline::XlsxOutline, scheme::color_scheme::XlsxColorScheme, st_types::emu_to_pt,
};

use super::{
    compound_line::CompoundLineValues, join_type::LineJoinTypeValue, line_cap::LineCapValues,
    line_dash::LineDashTypeValues, line_end::LineEnd,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outline?view=openxml-3.0.1
///
/// This element specifies an outline style that can be applied to a number of different objects such as shapes and text.
/// The line allows for the specifying of many different types of outlines including even line dashes and bevels.
///
/// Example:
/// ```
/// <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
///     <a:solidFill>
///         <a:schemeClr val="phClr">
///             <a:shade val="95000" />
///             <a:satMod val="104999" />
///         </a:schemeClr>
///     </a:solidFill>
///     <a:prstDash val="solid" />
/// </a:ln>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct Outline {
    pub line_join: LineJoinTypeValue,

    pub dash: LineDashTypeValues,

    pub fill: Fill,

    /// headEnd (Line Head/End Style)
    ///
    /// This element specifies decorations which can be added to the head of a line.
    ///
    /// Example:
    /// ```
    /// <headEnd len="lg" type="arrowhead" w="sm"/>
    /// ```
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.headend?view=openxml-3.0.1
    pub line_head: LineEnd,

    /// tailEnd (Tail line end style)
    ///
    /// This element specifies decorations which can be added to the tail of a line.
    ///
    /// Example
    /// ```
    /// <tailEnd len="lg" type="arrowhead" w="sm"/>
    /// ```
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tailend?view=openxml-3.0.1
    pub line_tail: LineEnd,

    /// algn (Stroke Alignment)
    ///
    /// Specifies the alignment to be used for the underline stroke.
    ///
    /// * Center
    /// * Insert
    pub stroke_alignment: PenAlignmentValues,

    /// cap (Line Ending Cap Type)
    ///
    /// Specifies the ending caps that should be used for this line such as rounded, flat, etc
    ///
    /// * Flat
    /// * Round
    /// * Square
    pub line_cap: LineCapValues,

    /// cmpd (Compound Line Type)
    ///
    /// Specifies the compound line type to be used for the underline stroke.
    ///
    /// * Double
    /// * Single
    /// * ThickThin
    /// * ThinThick
    /// * Triple
    pub compound_line_type: CompoundLineValues,

    /// w (Line Width)
    ///
    /// Specifies the width to be used for the underline stroke.
    /// default to 1.5pt (12700 emu)
    pub width: f64,
}

impl Outline {
    pub(crate) fn from_raw(
        raw: Option<XlsxOutline>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };

        return Some(Self {
            line_join: LineJoinTypeValue::from_raw(raw.clone(), None),
            dash: LineDashTypeValues::from_raw(raw.clone(), None),
            fill: Fill::from_outline(raw.clone(), color_scheme.clone(), ref_color),
            line_head: LineEnd::from_raw(raw.clone().head_end, None),
            line_tail: LineEnd::from_raw(raw.clone().tail_end, None),
            stroke_alignment: PenAlignmentValues::from_string(raw.clone().alignment, None),
            line_cap: LineCapValues::from_string(raw.clone().cap, None),
            compound_line_type: CompoundLineValues::from_string(raw.clone().compound, None),
            width: emu_to_pt(raw.clone().w.unwrap_or(19050) as i64),
        });
    }

    pub(crate) fn with_reference(
        raw: Option<XlsxOutline>,
        line_ref: Option<Outline>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Option<Self> {
        let line_ref = if let Some(l) = line_ref.clone() {
            let mut l = l;
            l.width = 1.5;
            Some(l)
        } else {
            None
        };

        let Some(raw) = raw else { return line_ref };

        let fill = if raw.clone().no_fill.is_some()
            || raw.clone().gradient_fill.is_some()
            || raw.clone().solid_fill.is_some()
            || raw.clone().pattern_fill.is_some()
        {
            Fill::from_outline(raw.clone(), color_scheme.clone(), None)
        } else if let Some(r) = line_ref.clone() {
            r.fill
        } else {
            Fill::SolidFill("020b0fff".to_string())
        };

        // width is disrespecting the refernece
        let width = emu_to_pt(raw.clone().w.unwrap_or(19050) as i64);

        return Some(Self {
            line_join: LineJoinTypeValue::from_raw(
                raw.clone(),
                if let Some(r) = line_ref.clone() {
                    Some(r.line_join)
                } else {
                    None
                },
            ),
            dash: LineDashTypeValues::from_raw(
                raw.clone(),
                if let Some(r) = line_ref.clone() {
                    Some(r.dash)
                } else {
                    None
                },
            ),
            fill,
            line_head: LineEnd::from_raw(
                raw.clone().head_end,
                if let Some(r) = line_ref.clone() {
                    Some(r.line_head)
                } else {
                    None
                },
            ),
            line_tail: LineEnd::from_raw(
                raw.clone().tail_end,
                if let Some(r) = line_ref.clone() {
                    Some(r.line_tail)
                } else {
                    None
                },
            ),
            stroke_alignment: PenAlignmentValues::from_string(
                raw.clone().alignment,
                if let Some(r) = line_ref.clone() {
                    Some(r.stroke_alignment)
                } else {
                    None
                },
            ),
            line_cap: LineCapValues::from_string(
                raw.clone().cap,
                if let Some(r) = line_ref.clone() {
                    Some(r.line_cap)
                } else {
                    None
                },
            ),
            compound_line_type: CompoundLineValues::from_string(
                raw.clone().compound,
                if let Some(r) = line_ref.clone() {
                    Some(r.compound_line_type)
                } else {
                    None
                },
            ),
            width,
        });
    }
}
