use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::{
    excel::XmlReader,
    raw::drawing::{
        graphic::graphic_frame::XlsxGraphicFrame,
        image::picture::XlsxPicture,
        non_visual_properties::non_visual_group_shape_properties::XlsxNonVisualGroupShapeProperties,
        shape::{
            connection_shape::XlsxConnectionShape,
            visual_group_shape_properties::XlsxVisualGroupShapeProperties,
        },
        worksheet_drawing::{
            drawing_content_type::XlsxWorksheetDrawingContentType, spreadsheet_shape::XlsxShape,
        },
    },
};

/// - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.groupshape?view=openxml-3.0.1
///
/// This element specifies a group shape that represents many shapes grouped together.
/// This shape is to be treated just as if it were a regular shape but instead of being described by a single geometry it is made up of all the shape geometries encompassed within it.
/// Within a group shape each of the shapes that make up the group are specified just as they normally would.
/// The idea behind grouping elements however is that a single transform can apply to many shapes at the same time.
///
/// Example:
///
/// ```
/// <p:grpSp>
///   <p:nvGrpSpPr>
///     <p:cNvPr id="10" name="Group 9"/>
///     <p:cNvGrpSpPr/>
///     <p:nvPr/>
///   </p:nvGrpSpPr>
///   <p:grpSpPr>
///     <a:xfrm>
///       <a:off x="838200" y="990600"/>
///       <a:ext cx="2426208" cy="978408"/>
///       <a:chOff x="838200" y="990600"/>
///       <a:chExt cx="2426208" cy="978408"/>
///     </a:xfrm>
///   </p:grpSpPr>
///   <p:sp>
///   …  </p:sp>
///   <p:sp>
///   …  </p:sp>
///   <p:sp>
///   …  </p:sp>
/// </p:grpSp>
/// ```
///
/// grpSp (Group shape)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxGroupShape {
    // Child Elements
    // extLst (Extension List) Not supported
    /// grpSpPr (Visual Group Shape Properties)	§20.1.2.2.22
    pub visual_group_shape_properties: Option<XlsxVisualGroupShapeProperties>,

    /// nvGrpSpPr (Non-Visual Properties for a Group Shape)	§20.1.2.2.27
    pub non_visual_group_shape_properties: Option<XlsxNonVisualGroupShapeProperties>,

    /// group of the following
    /// - pic (Picture)
    /// - sp (Shape)
    /// - grpSp (Group shape)
    /// - graphicFrame (Graphic Frame) (Ex: Charts)
    /// - cxnSp (Connection Shape)	§20.1.2.2.10
    ///
    /// (content parts cannot appear in a group shape)
    pub drawing_contents: Option<Box<Vec<XlsxWorksheetDrawingContentType>>>,
}

impl XlsxGroupShape {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut shape = Self {
            visual_group_shape_properties: None,
            non_visual_group_shape_properties: None,
            drawing_contents: None,
        };

        let mut shapes: Vec<XlsxWorksheetDrawingContentType> = vec![];

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cxnSp" => {
                    shapes.push(XlsxWorksheetDrawingContentType::ConnectionShape(
                        XlsxConnectionShape::load(reader, e)?,
                    ));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpSp" => {
                    shapes.push(XlsxWorksheetDrawingContentType::GroupShape(
                        XlsxGroupShape::load(reader)?,
                    ));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpSpPr" => {
                    shape.visual_group_shape_properties =
                        Some(XlsxVisualGroupShapeProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"graphicFrame" => {
                    shapes.push(XlsxWorksheetDrawingContentType::GraphicFrame(
                        XlsxGraphicFrame::load(reader, e)?,
                    ));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"nvGrpSpPr" => {
                    shape.non_visual_group_shape_properties =
                        Some(XlsxNonVisualGroupShapeProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pic" => {
                    shapes.push(XlsxWorksheetDrawingContentType::Picture(XlsxPicture::load(
                        reader, e,
                    )?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sp" => {
                    shapes.push(XlsxWorksheetDrawingContentType::Shape(XlsxShape::load(
                        reader, e,
                    )?));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"grpSp" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `grpSp`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        shape.drawing_contents = Some(Box::new(shapes));

        return Ok(shape);
    }
}
