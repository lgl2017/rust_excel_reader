use crate::raw::drawing::{
    graphic::graphic_frame::XlsxGraphicFrame, image::picture::XlsxPicture,
    shape::connection_shape::XlsxConnectionShape,
};

use super::{group_shape::XlsxGroupShape, spreadsheet_shape::XlsxShape};

/// enum for the following types
/// - pic (Picture)
/// - sp (Shape)
/// - txSp (Text Shape)
/// - grpSp (Group shape)
/// - graphic frame (Graphic Frame) (Ex: Charts)
#[derive(Debug, Clone, PartialEq)]
pub enum XlsxWorksheetDrawingContentType {
    /// grpSp (Group Shape)	§20.5.2.17
    GroupShape(XlsxGroupShape),

    /// cxnSp (Connection Shape)	§20.5.2.13
    ConnectionShape(XlsxConnectionShape),

    /// sp (Shape)	§20.5.2.29
    Shape(XlsxShape),

    /// pic (Picture)	§20.5.2.25
    Picture(XlsxPicture),

    // graphicFrame (Graphic Frame) (Ex: Charts)
    GraphicFrame(XlsxGraphicFrame),
}
