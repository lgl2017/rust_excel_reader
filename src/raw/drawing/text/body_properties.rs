use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_int, string_to_unsignedint},
    raw::drawing::{
        scene::scene_3d_type::XlsxScene3DType,
        shape::{shape_3d_type::XlsxShape3DType, shape_autofit::XlsxShapeAutofit},
        st_types::{STAngle, STCoordinate, STPositiveCoordinate},
    },
};

use super::{
    flat_text::XlsxFlatText, no_autofit::XlsxNoAutoFit, norm_autofit::XlsxNormAutoFit,
    preset_text_warp::XlsxPresetTextWarp,
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
pub struct XlsxBodyProperties {
    // child: extLst (Extension List)	Not supported

    // Child Elements
    // flatTx (No text in 3D scene)	§20.1.5.8
    pub flat_tx: Option<XlsxFlatText>,

    /// NoAutoFit:  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.noautofit?view=openxml-3.0.1
    /// This element specifies that text within the text body should not be auto-fit to the bounding box.
    /// Auto-fitting is when text within a text box is scaled in order to remain inside the text box.
    ///
    /// Example
    /// ```
    /// <a:bodyPr wrap="none" rtlCol="0">
    ///     <a:noAutofit/>
    /// </a:bodyPr>
    /// ```
    // noAutofit (No AutoFit)	§21.1.2.1.2
    pub no_autofit: Option<XlsxNoAutoFit>,

    /// NormalAutoFit: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.normalautofit?view=openxml-3.0.1
    /// This element specifies that text within the text body should be normally auto-fit to the bounding box.
    /// If this element is omitted, then noAutofit or auto-fit off is implied.
    ///
    /// Example
    /// ```
    /// <a:bodyPr wrap="none" rtlCol="0">
    ///     <a:normAutofit fontScale="92.000%" lnSpcReduction="20.000%"/>
    /// </a:bodyPr>
    /// ```
    // normAutofit (Normal AutoFit)	§21.1.2.1.3
    pub norm_autofit: Option<XlsxNormAutoFit>,

    // prstTxWarp (Preset Text Warp)	§20.1.9.19
    pub preset_text_warp: Option<XlsxPresetTextWarp>,

    // scene3d (3D Scene Properties)	§20.1.4.1.26
    pub scene3d: Option<XlsxScene3DType>,

    // sp3d (Apply 3D shape properties)	§20.1.5.12
    pub shape3d: Option<XlsxShape3DType>,

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
    // spAutoFit (Shape AutoFit)
    pub shape_autofit: Option<XlsxShapeAutofit>,

    // attributes
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.anchor?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-anchor
    ///
    /// Specifies the anchoring position of the txBody within the shape.
    /// If this attribute is omitted, then a value of t, meaning top, is implied.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textanchoringtypevalues?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <a:bodyPr anchor="ctr" … />
    /// ```
    // tag: anchor
    pub anchor: Option<String>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.anchorcenter?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-anchorcenter
    ///
    /// Specifies the centering of the text box.
    /// If this attribute is omitted, then a value of 0 (False) is implied.
    ///
    /// Example:
    /// ```
    /// <a:bodyPr anchor="ctr" anchorCtr="1" … />
    /// ```
    // tag: anchorCtr
    pub anchor_center: Option<bool>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.bottominset?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-bottominset
    ///
    /// Specifies the bottom inset of the bounding rectangle.
    /// If this attribute is omitted, a value of 45720, meaning 0.05 inches, is implied.
    ///
    /// Example
    /// ```
    /// <a:bodyPr lIns="91440" tIns="91440" rIns="91440" bIns="91440" … />
    /// ```
    ///
    // tag: bIns
    pub bottom_inset: Option<STCoordinate>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.compatiblelinespacing?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-compatiblelinespacing
    ///
    /// Specifies that the line spacing for this text body will be decided in a simplistic manner using the font scene.
    /// If this attribute is omitted, a value of 0, meaning false, is implied.
    // tag: compatLnSpc
    pub compatible_line_spacing: Option<bool>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.forceantialias?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-forceantialias
    ///
    /// Forces the text to be rendered anti-aliased regardless of the font size.
    /// If this attribute is omitted, then a value of 0, meaning off, is implied.
    // tag: forceAA
    pub force_anti_alias: Option<bool>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.fromwordart?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-fromwordart
    ///
    /// Specifies that text within this textbox is converted text from a WordArt object.
    /// If this attribute is omitted, then a value of 0, meaning off, is implied.
    // tag: fromWordArt
    pub from_word_art: Option<bool>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.horizontaloverflow?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-horizontaloverflow
    ///
    ///  Determines whether the text can flow out of the bounding box horizontally.
    /// This is used to determine what will happen in the event that the text within a shape is too large for the bounding box it is contained within.
    /// If this attribute is omitted, then a value of `overflow` is implied.
    ///
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.texthorizontaloverflowvalues?view=openxml-3.0.1
    // tag: horzOverflow
    pub horizontal_overflow: Option<String>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.leftinset?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-leftinset
    ///
    /// Specifies the left inset of the bounding rectangle.
    /// If this attribute is omitted, then a value of 91440, meaning 0.1 inches, is implied.
    // tag: lIns
    pub left_inset: Option<STCoordinate>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.columncount?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-columncount
    ///
    /// Specifies the number of columns of text in the bounding rectangle.
    /// If this attribute is omitted, then a value of 1 is implied.
    // tag: numCol
    pub column_count: Option<u64>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.rightinset?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-rightinset
    ///
    /// Specifies the right inset of the bounding rectangle.
    /// If this attribute is omitted, then a value of 91440, meaning 0.1 inches, is implied.
    // tag: rIns
    pub right_inset: Option<STCoordinate>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.rotation?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-rotation
    ///
    /// Specifies the rotation that is being applied to the text within the bounding box.
    /// if this attribute is omitted, then a value of 0 is implied.
    // tag: rot
    pub rotation: Option<STAngle>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.righttoleftcolumns?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-righttoleftcolumns
    ///
    /// Specifies whether columns are used in a right-to-left or left-to-right order.
    /// If this attribute is omitted, then a value of 0 is implied in which case text will start in the leftmost column and flow to the right.
    // tag: rtlCol
    pub right_to_left_columns: Option<bool>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.columnspacing?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-columnspacing
    ///
    /// Specifies the space between text columns in the text area.
    ///  If this attribute is omitted, then a value of 0 is implied.
    // tag: spcCol
    pub column_spacing: Option<STPositiveCoordinate>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.useparagraphspacing?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-useparagraphspacing
    ///
    /// Specifies whether the before and after paragraph spacing defined by the user is to be respected.
    /// If this attribute is omitted, then a value of 0, meaning false, is implied.
    // tag: spcFirstLastPara
    pub use_paragraph_spacing: Option<bool>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.topinset?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-topinset
    ///
    /// Specifies the top inset of the bounding rectangle.
    /// If this attribute is omitted, then a value of 45720, meaning 0.05 inches, is implied.
    // tag: tIns
    pub top_inset: Option<STCoordinate>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.upright?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-upright
    ///
    /// Specifies whether text should remain upright, regardless of the transform applied to it and the accompanying shape transform.
    /// If this attribute is omitted, then a value of 0, meaning false, will be implied.
    // tag: upright
    pub upright: Option<bool>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.vertical?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-vertical
    ///
    /// Determines if the text within the given text body should be displayed vertically.
    /// If this attribute is omitted, then a value of `horz`, meaning no vertical text, is implied.
    ///
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textverticalvalues?view=openxml-3.0.1
    // tag: vert
    pub vertical: Option<String>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.verticaloverflow?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-verticaloverflow
    ///
    /// Determines whether the text can flow out of the bounding box vertically.
    ///  If this attribute is omitted, then a value of `overflow` is implied.
    ///
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textverticaloverflowvalues?view=openxml-3.0.1
    // tag: vertOverflow
    pub vertical_overflow: Option<String>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.wrap?view=openxml-3.0.1#documentformat-openxml-drawing-bodyproperties-wrap
    ///
    /// Specifies the wrapping options to be used for this text body.
    /// If this attribute is omitted, then a value of `square` is implied which will wrap the text using the bounding text box.
    ///
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textwrappingvalues?view=openxml-3.0.1
    // tag: wrap
    pub wrap: Option<String>,
}

impl XlsxBodyProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut properties = Self {
            flat_tx: None,
            no_autofit: None,
            norm_autofit: None,
            preset_text_warp: None,
            scene3d: None,
            shape3d: None,
            shape_autofit: None,
            anchor: None,
            anchor_center: None,
            bottom_inset: None,
            compatible_line_spacing: None,
            force_anti_alias: None,
            from_word_art: None,
            horizontal_overflow: None,
            left_inset: None,
            column_count: None,
            right_inset: None,
            rotation: None,
            right_to_left_columns: None,
            column_spacing: None,
            use_paragraph_spacing: None,
            top_inset: None,
            upright: None,
            vertical: None,
            vertical_overflow: None,
            wrap: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"anchor" => properties.anchor = Some(string_value),
                        b"anchorCtr" => properties.anchor_center = string_to_bool(&string_value),
                        b"bIns" => properties.bottom_inset = string_to_int(&string_value),
                        b"compatLnSpc" => {
                            properties.compatible_line_spacing = string_to_bool(&string_value)
                        }
                        b"forceAA" => properties.force_anti_alias = string_to_bool(&string_value),
                        b"fromWordArt" => properties.from_word_art = string_to_bool(&string_value),
                        b"horzOverflow" => properties.horizontal_overflow = Some(string_value),
                        b"lIns" => properties.left_inset = string_to_int(&string_value),
                        b"numCol" => properties.column_count = string_to_unsignedint(&string_value),
                        b"rIns" => properties.right_inset = string_to_int(&string_value),
                        b"rot" => properties.rotation = string_to_int(&string_value),
                        b"rtlCol" => {
                            properties.right_to_left_columns = string_to_bool(&string_value)
                        }
                        b"spcCol" => {
                            properties.column_spacing = string_to_unsignedint(&string_value)
                        }
                        b"spcFirstLastPara" => {
                            properties.use_paragraph_spacing = string_to_bool(&string_value)
                        }
                        b"tIns" => properties.top_inset = string_to_int(&string_value),
                        b"upright" => properties.upright = string_to_bool(&string_value),
                        b"vert" => properties.vertical = Some(string_value),
                        b"vertOverflow" => properties.vertical_overflow = Some(string_value),
                        b"wrap" => properties.wrap = Some(string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"flatTx" => {
                    properties.flat_tx = Some(XlsxFlatText::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"noAutofit" => {
                    properties.no_autofit = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"normAutofit" => {
                    properties.norm_autofit = Some(XlsxNormAutoFit::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"prstTxWarp" => {
                    properties.preset_text_warp = Some(XlsxPresetTextWarp::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"scene3d" => {
                    properties.scene3d = Some(XlsxScene3DType::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sp3d" => {
                    properties.shape3d = Some(XlsxShape3DType::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spAutoFit" => {
                    properties.shape_autofit = Some(true);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"bodyPr" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        if properties.norm_autofit.is_none() && properties.shape_autofit.is_none() {
            properties.no_autofit = Some(true);
        }

        Ok(properties)
    }
}
