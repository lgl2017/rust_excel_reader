pub mod absolute_anchor;
pub mod client_data;
pub mod content_part;
pub mod drawing_content_type;
pub mod group_shape;
pub mod marker;
pub mod one_cell_anchor;
pub mod spreadsheet_extent;
pub mod spreadsheet_position;
pub mod spreadsheet_shape;
pub mod two_cell_anchor;

use absolute_anchor::XlsxAbsoluteAnchor;
use anyhow::bail;
use one_cell_anchor::XlsxOneCellAnchor;
use quick_xml::events::Event;
use std::io::{Read, Seek};

use two_cell_anchor::XlsxTwoCellAnchor;
use zip::ZipArchive;

use crate::excel::xml_reader;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.worksheetdrawing?view=openxml-3.0.1
///
/// This element specifies all drawing objects within the worksheet.
///
/// Example :
///
/// ```
/// <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
///     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
///     <xdr:twoCellAnchor editAs="oneCell">
///         <xdr:from>
///             <xdr:col>2</xdr:col>
///             <xdr:colOff>825500</xdr:colOff>
///             <xdr:row>2</xdr:row>
///             <xdr:rowOff>241300</xdr:rowOff>
///         </xdr:from>
///         <xdr:to>
///             <xdr:col>5</xdr:col>
///             <xdr:colOff>901700</xdr:colOff>
///             <xdr:row>17</xdr:row>
///             <xdr:rowOff>38100</xdr:rowOff>
///         </xdr:to>
///         <xdr:pic>
///             <xdr:nvPicPr>
///                 <xdr:cNvPr id="2" name="Picture 1">
///                     <a:extLst>
///                         <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
///                             <a16:creationId
///                                 xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"
///                                 id="{747A616F-59EB-8B49-A6F6-BF46B404B319}" />
///                         </a:ext>
///                     </a:extLst>
///                 </xdr:cNvPr>
///                 <xdr:cNvPicPr>
///                     <a:picLocks noChangeAspect="1" />
///                 </xdr:cNvPicPr>
///             </xdr:nvPicPr>
///             <xdr:blipFill>
///                 <a:blip
///                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
///                     r:embed="rId1" />
///                 <a:stretch>
///                     <a:fillRect />
///                 </a:stretch>
///             </xdr:blipFill>
///             <xdr:spPr>
///                 <a:xfrm>
///                     <a:off x="3314700" y="838200" />
///                     <a:ext cx="3810000" cy="3581400" />
///                 </a:xfrm>
///                 <a:prstGeom prst="rect">
///                     <a:avLst />
///                 </a:prstGeom>
///             </xdr:spPr>
///         </xdr:pic>
///         <xdr:clientData />
///     </xdr:twoCellAnchor>
///     <xdr:twoCellAnchor editAs="oneCell">
///         <xdr:from>
///             <xdr:col>6</xdr:col>
///             <xdr:colOff>190500</xdr:colOff>
///             <xdr:row>14</xdr:row>
///             <xdr:rowOff>203200</xdr:rowOff>
///         </xdr:from>
///         <xdr:to>
///             <xdr:col>9</xdr:col>
///             <xdr:colOff>266700</xdr:colOff>
///             <xdr:row>28</xdr:row>
///             <xdr:rowOff>228600</xdr:rowOff>
///         </xdr:to>
///         <xdr:pic>
///             <xdr:nvPicPr>
///                 <xdr:cNvPr id="3" name="Picture 2">
///                     <a:extLst>
///                         <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
///                             <a16:creationId
///                                 xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"
///                                 id="{C18AFA89-ECBE-E240-835F-EF9A9F44A182}" />
///                         </a:ext>
///                     </a:extLst>
///                 </xdr:cNvPr>
///                 <xdr:cNvPicPr>
///                     <a:picLocks noChangeAspect="1" />
///                 </xdr:cNvPicPr>
///             </xdr:nvPicPr>
///             <xdr:blipFill>
///                 <a:blip
///                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
///                     r:embed="rId1" />
///                 <a:stretch>
///                     <a:fillRect />
///                 </a:stretch>
///             </xdr:blipFill>
///             <xdr:spPr>
///                 <a:xfrm>
///                     <a:off x="7658100" y="3822700" />
///                     <a:ext cx="3810000" cy="3581400" />
///                 </a:xfrm>
///                 <a:prstGeom prst="rect">
///                     <a:avLst />
///                 </a:prstGeom>
///             </xdr:spPr>
///         </xdr:pic>
///         <xdr:clientData />
///     </xdr:twoCellAnchor>
/// </xdr:wsDr>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxWorksheetDrawing {
    // Child Elements
    // absoluteAnchor (Absolute Anchor Shape Size)	§20.5.2.1
    // oneCellAnchor (One Cell Anchor Shape Size)	§20.5.2.24
    // twoCellAnchor (Two Cell Anchor Shape Size)	§20.5.2.33
    pub drawings: Option<Vec<XlsxWorksheetDrawingType>>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum XlsxWorksheetDrawingType {
    // absoluteAnchor (Absolute Anchor Shape Size)
    AbsoluteAnchor(XlsxAbsoluteAnchor),

    // oneCellAnchor (One Cell Anchor Shape Size)	§20.5.2.24
    OneCellAnchor(XlsxOneCellAnchor),

    // twoCellAnchor (Two Cell Anchor Shape Size)	§20.5.2.33
    TwoCellAnchor(XlsxTwoCellAnchor),
}

impl XlsxWorksheetDrawing {
    pub(crate) fn load(zip: &mut ZipArchive<impl Read + Seek>, path: &str) -> anyhow::Result<Self> {
        let mut worksheet_drawing = Self { drawings: None };
        let Some(mut reader) = xml_reader(zip, path) else {
            return Ok(worksheet_drawing);
        };

        let mut drawings: Vec<XlsxWorksheetDrawingType> = vec![];

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"absoluteAnchor" => {
                    let drawing = XlsxAbsoluteAnchor::load(&mut reader)?;
                    drawings.push(XlsxWorksheetDrawingType::AbsoluteAnchor(drawing));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"oneCellAnchor" => {
                    let drawing = XlsxOneCellAnchor::load(&mut reader)?;
                    drawings.push(XlsxWorksheetDrawingType::OneCellAnchor(drawing));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"twoCellAnchor" => {
                    let drawing = XlsxTwoCellAnchor::load(&mut reader, e)?;
                    drawings.push(XlsxWorksheetDrawingType::TwoCellAnchor(drawing));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"wsDr" => break,
                Ok(Event::Eof) => break,
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        worksheet_drawing.drawings = Some(drawings);
        Ok(worksheet_drawing)
    }
}
