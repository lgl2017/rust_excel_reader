#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    processed::drawing::{
        scene::scene_3d_properties::Scene3DProperties,
        shape::shape_3d_properties::Shape3DProperties,
    },
    raw::drawing::{
        scheme::color_scheme::XlsxColorScheme,
        st_types::{emu_to_pt, st_angle_to_degree},
        text::body_properties::XlsxBodyProperties,
    },
};

use super::{
    horizontal_overflow::TextHorizontalOverflowValues, normal_auto_fit::NormalAutoFit,
    preset_text_shape::PresetTextShapeValues, text_anchoring_type::TextAnchoringTypeValues,
    text_vertical_values::TextVerticalValues, text_wrapping::TextWrappingValues,
    vertical_overflow::TextVerticalOverflowValues,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties?view=openxml-3.0.1
///
/// defines the body properties for the text body within a shape.
///
/// Example:
/// ```
/// <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="50800" tIns="50800" rIns="50800" bIns="50800" numCol="1" spcCol="38100" rtlCol="0" anchor="ctr" upright="0">
///     <a:spAutoFit />
///     <prstTxWarp prst="textArchUp">
///         <a:avLst>
///             <a:gd name="myGuide" fmla="val 2"/>
///         </a:avLst>
///     </prstTxWarp>
/// </a:bodyPr>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct BodyProperties {
    /// auto fit behavior:
    pub auto_fit: AutoFitTypeValues,

    /// prstTxWarp (Preset Text Warp)
    pub preset_text_warp: PresetTextShapeValues,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.anchor?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-anchor
    ///
    /// Specifies the anchoring position of the txBody within the shape.
    /// * Bottom
    /// * Center
    /// * Top
    ///
    /// If this attribute is omitted, then a value of t, meaning top, is implied.
    ///
    /// Example:
    /// ```
    /// <a:bodyPr anchor="ctr" … />
    /// ```
    pub anchor: TextAnchoringTypeValues,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.anchorcenter?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-anchorcenter
    ///
    /// Specifies the centering of the text box.
    /// If this attribute is omitted, then a value of 0 (False) is implied.
    ///
    /// Example:
    /// ```
    /// <a:bodyPr anchor="ctr" anchorCtr="1" … />
    /// ```
    pub anchor_center: bool,

    /// Specifies the insets of the bounding rectangle.
    pub insets: BoundingRectInsets,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.vertical?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-vertical
    ///
    /// Determines if the text within the given text body should be displayed vertically.
    /// If this attribute is omitted, then a value of `horz`, meaning no vertical text, is implied.
    ///
    /// * EastAsianVetical
    /// * Horizontal
    /// * MongolianVertical
    /// * Vertical
    /// * Vertical270
    /// * WordArtLeftToRight
    /// * WordArtVertical
    pub text_display: TextVerticalValues,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.wrap?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-wrap
    ///
    /// Specifies the wrapping options to be used for this text body.
    /// If this attribute is omitted, then a value of `square` is implied which will wrap the text using the bounding text box.
    ///
    /// * None
    /// * Square
    pub text_wrap: TextWrappingValues,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.horizontaloverflow?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-horizontaloverflow
    ///
    /// Determines whether the text can flow out of the bounding box horizontally.
    ///
    /// This is used to determine what will happen in the event that the text within a shape is too large for the bounding box it is contained within.
    /// If this attribute is omitted, then a value of `overflow` is implied.
    ///
    /// * Clip
    /// * Overflow
    pub horizontal_overflow: TextHorizontalOverflowValues,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.verticaloverflow?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-verticaloverflow
    ///
    /// Determines whether the text can flow out of the bounding box vertically.
    /// If this attribute is omitted, then a value of `overflow` is implied.
    ///
    /// * Clip
    /// * Ellipse
    /// * Overflow
    pub vertical_overflow: TextVerticalOverflowValues,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.columncount?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-columncount
    ///
    /// Specifies the number of columns of text in the bounding rectangle.
    /// If this attribute is omitted, then a value of 1 is implied.
    ///
    /// value range:
    /// - minimum value of greater than or equal to 1.
    /// - maximum value of less than or equal to 16.
    pub column_count: u64,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.columnspacing?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-columnspacing
    ///
    /// Specifies the space between text columns in the text area.
    /// If this attribute is omitted, then a value of 0 is implied.
    pub column_spacing: f64,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.rotation?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-rotation
    ///
    /// Specifies the rotation that is being applied to the text within the bounding box.
    /// if this attribute is omitted, then a value of 0  is implied.
    pub rotation: f64,

    /// flatTx (No text in 3D scene)
    ///
    /// Specifies the Z coordinate to be used in positioning the flat text within the 3D scene.
    pub flat_text_z: f64,

    /// scene3d (3D Scene Properties)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub scene3d_properties: Option<Scene3DProperties>,

    /// sp3d (Apply 3D shape properties)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub shape3d_properties: Option<Shape3DProperties>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.compatiblelinespacing?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-compatiblelinespacing
    ///
    /// Specifies that the line spacing for this text body will be decided in a simplistic manner using the font scene.
    /// If this attribute is omitted, a value of 0, meaning false, is implied.
    pub compatible_line_spacing: bool,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.forceantialias?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-forceantialias
    ///
    /// Forces the text to be rendered anti-aliased regardless of the font size.
    /// If this attribute is omitted, then a value of 0, meaning off, is implied.
    pub force_anti_alias: bool,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.fromwordart?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-fromwordart
    ///
    /// Specifies that text within this textbox is converted text from a WordArt object.
    /// If this attribute is omitted, then a value of 0, meaning off, is implied.
    pub from_word_art: bool,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.righttoleftcolumns?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-righttoleftcolumns
    ///
    /// Specifies whether columns are used in a right-to-left or left-to-right order.
    /// If this attribute is omitted, then a value of 0 is implied in which case text will start in the leftmost column and flow to the right.
    pub right_to_left_columns: bool,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.useparagraphspacing?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-useparagraphspacing
    ///
    /// Specifies whether the before and after paragraph spacing defined by the user is to be respected.
    /// If this attribute is omitted, then a value of 0, meaning false, is implied.
    pub use_paragraph_spacing: bool,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.upright?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-upright
    ///
    /// Specifies whether text should remain upright, regardless of the transform applied to it and the accompanying shape transform.
    /// If this attribute is omitted, then a value of 0, meaning false, will be implied.
    pub upright: bool,
}

impl BodyProperties {
    pub(crate) fn default() -> Self {
        Self {
            auto_fit: AutoFitTypeValues::default(),
            preset_text_warp: PresetTextShapeValues::default(),
            anchor: TextAnchoringTypeValues::default(),
            anchor_center: false,
            insets: BoundingRectInsets::default(),
            text_display: TextVerticalValues::default(),
            text_wrap: TextWrappingValues::default(),
            horizontal_overflow: TextHorizontalOverflowValues::default(),
            vertical_overflow: TextVerticalOverflowValues::default(),
            column_count: 1,
            column_spacing: 0.0,
            rotation: 0.0,
            flat_text_z: 0.0,
            scene3d_properties: None,
            shape3d_properties: None,
            compatible_line_spacing: false,
            force_anti_alias: false,
            from_word_art: false,
            right_to_left_columns: false,
            use_paragraph_spacing: false,
            upright: false,
        }
    }

    pub(crate) fn from_raw(
        raw: Option<XlsxBodyProperties>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };
        let preset = if let Some(p) = raw.clone().preset_text_warp {
            PresetTextShapeValues::from_string(p.preset)
        } else {
            PresetTextShapeValues::default()
        };

        let column_count = if let Some(c) = raw.clone().column_count {
            if (1..=16).contains(&c) {
                c
            } else {
                1
            }
        } else {
            1
        };

        let flat_text_z = if let Some(ft) = raw.clone().flat_tx {
            emu_to_pt(ft.z.unwrap_or(0))
        } else {
            0.0
        };

        return Self {
            auto_fit: AutoFitTypeValues::from_raw(raw.clone()),
            preset_text_warp: preset,
            anchor: TextAnchoringTypeValues::from_string(raw.clone().anchor),
            anchor_center: false,
            insets: BoundingRectInsets::from_body_properties(raw.clone()),
            text_display: TextVerticalValues::from_string(raw.clone().vertical),
            text_wrap: TextWrappingValues::from_string(raw.clone().wrap),
            horizontal_overflow: TextHorizontalOverflowValues::from_string(
                raw.clone().horizontal_overflow,
            ),
            vertical_overflow: TextVerticalOverflowValues::from_string(
                raw.clone().vertical_overflow,
            ),
            column_count,
            column_spacing: emu_to_pt(raw.clone().column_spacing.unwrap_or(0) as i64),
            rotation: st_angle_to_degree(raw.clone().rotation.unwrap_or(0)),
            flat_text_z,
            scene3d_properties: Scene3DProperties::from_raw(raw.clone().scene3d),
            shape3d_properties: Shape3DProperties::from_raw(
                raw.clone().shape3d,
                color_scheme.clone(),
                None, // reference color for text body is used for font only
            ),
            compatible_line_spacing: raw.compatible_line_spacing.unwrap_or(false),
            force_anti_alias: raw.force_anti_alias.unwrap_or(false),
            from_word_art: raw.from_word_art.unwrap_or(false),
            right_to_left_columns: raw.right_to_left_columns.unwrap_or(false),
            use_paragraph_spacing: raw.use_paragraph_spacing.unwrap_or(false),
            upright: raw.upright.unwrap_or(false),
        };
    }
}

/// * ShapeAutoFit
/// * NoAutoFit
/// * NormalAutoFit
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum AutoFitTypeValues {
    /// Shape AutoFit:
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeautofit?view=openxml-3.0.1
    ///
    /// This element specifies that a shape should be auto-fit to fully contain the text described within it.
    /// Auto-fitting is when text within a shape is scaled in order to contain all the text inside.
    /// If this element is omitted, then noAutofit or auto-fit off is implied.
    ///
    /// Example:
    /// ```
    /// <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="50800" tIns="50800" rIns="50800" bIns="50800" numCol="1" spcCol="38100" rtlCol="0" anchor="ctr" upright="0">
    ///     <a:spAutoFit />
    /// </a:bodyPr>
    /// ```
    ShapeAutoFit,

    /// NoAutoFit
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.noautofit?view=openxml-3.0.1
    ///
    /// This element specifies that text within the text body should not be auto-fit to the bounding box.
    /// Auto-fitting is when text within a text box is scaled in order to remain inside the text box.
    ///
    /// Example
    /// ```
    /// <a:bodyPr wrap="none" rtlCol="0">
    ///     <a:noAutofit/>
    /// </a:bodyPr>
    /// ```
    NoAutoFit,

    /// NormalAutoFit:
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.normalautofit?view=openxml-3.0.1
    ///
    /// This element specifies that text within the text body should be normally auto-fit to the bounding box.
    /// If this element is omitted, then noAutofit or auto-fit off is implied.
    ///
    /// Example
    /// ```
    /// <a:bodyPr wrap="none" rtlCol="0">
    ///     <a:normAutofit fontScale="92.000%" lnSpcReduction="20.000%"/>
    /// </a:bodyPr>
    /// ```
    NormalAutoFit(NormalAutoFit),
}

impl AutoFitTypeValues {
    pub(crate) fn default() -> Self {
        Self::NoAutoFit
    }

    pub(crate) fn from_raw(raw: XlsxBodyProperties) -> Self {
        if raw.shape_autofit == Some(true) {
            return Self::ShapeAutoFit;
        }
        if let Some(norm) = raw.norm_autofit {
            return Self::NormalAutoFit(NormalAutoFit::from_raw(norm));
        }
        return Self::default();
    }
}

/// Specifies the insets of the bounding rectangle
///
/// Example
/// ```
/// <a:bodyPr lIns="91440" tIns="91440" rIns="91440" bIns="91440" … />
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct BoundingRectInsets {
    /// Left Inset
    ///
    /// Specifies the left inset of the bounding rectangle.
    /// If this attribute is omitted, then a value of 91440, meaning 0.1 inches, is implied.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.leftinset?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-leftinset
    pub left: f64,

    /// Right Inset
    ///
    /// Specifies the right inset of the bounding rectangle.
    /// If this attribute is omitted, then a value of 91440, meaning 0.1 inches, is implied.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.rightinset?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-rightinset
    pub right: f64,

    /// Top inset
    ///
    /// Specifies the top inset of the bounding rectangle.
    /// If this attribute is omitted, then a value of 45720, meaning 0.05 inches, is implied.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.topinset?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-topinset
    pub top: f64,

    /// Bottom inset
    ///
    /// Specifies the bottom inset of the bounding rectangle.
    /// If this attribute is omitted, a value of 45720, meaning 0.05 inches, is implied.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.bottominset?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-bottominset
    pub bottom: f64,
}

impl BoundingRectInsets {
    pub(crate) fn default() -> Self {
        Self {
            left: emu_to_pt(91440),
            right: emu_to_pt(91440),
            top: emu_to_pt(45720),
            bottom: emu_to_pt(45720),
        }
    }

    pub(crate) fn from_body_properties(raw: XlsxBodyProperties) -> Self {
        return Self {
            left: emu_to_pt(raw.left_inset.clone().unwrap_or(91440)),
            right: emu_to_pt(raw.right_inset.clone().unwrap_or(91440)),
            top: emu_to_pt(raw.top_inset.clone().unwrap_or(45720)),
            bottom: emu_to_pt(raw.bottom_inset.clone().unwrap_or(45720)),
        };
    }
}
