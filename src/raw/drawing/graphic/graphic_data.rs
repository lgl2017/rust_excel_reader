use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.graphicdata?view=openxml-3.0.1
///
/// This element specifies the reference to a graphic object within the document.
/// This graphic object is provided entirely by the document authors who choose to persist this data within the document
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxGraphicData {
    // Child Elements
    // Any element in any namespace	n/a
    // memo: POI only processing if it is a chart.
    //
    // XSSFGraphicFrame.java:
    // if (graphicData != null) {
    //     NodeList nodes = graphicData.getDomNode().getChildNodes();
    //     for (int i = 0; i < nodes.getLength(); i++) {
    //         final Node node = nodes.item(i);
    //         // if the frame references a chart, associate the chart with this instance
    //         if (node.getAttributes() != null) {
    //             Node namedItem = node.getAttributes().getNamedItem("r:id");
    //             if (node.getNodeName().equals("c:chart") && namedItem != null) {
    //                 // this better succeed or the document is invalid
    //                 POIXMLDocumentPart relation = drawing.getRelationById(namedItem.getNodeValue());
    //                 // Do XWPF charts need similar treatment?
    //                 if (relation instanceof XSSFChart) {
    //                     ((XSSFChart) relation).setGraphicFrame(this);
    //                 }
    //             }
    //         }
    //     }
    // }
    //
    // sample xml:
    // <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
    //     <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    //         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    //         r:id="rId1" />
    // </a:graphicData>

    // Attributes
    /// uri (Uniform Resource Identifier)
    ///
    /// Specifies the uniform resource identifier that represents the data stored under this tag.
    /// The is used to identify the correct 'server' that can process the contents of this tag.
    pub uri: Option<String>,
}

impl XlsxGraphicData {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut graphic_data = Self { uri: None };

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"uri" => {
                            graphic_data.uri = Some(string_value);
                            break;
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(graphic_data)
    }
}
