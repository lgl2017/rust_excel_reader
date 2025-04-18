use std::io::{Read, Seek};

use anyhow::bail;
use calculation_properties::XlsxCalculationProperties;
use custom_workbook_view::{load_custom_bookviews, XlsxCustomWorkbookViews};
use defined_name::{load_defined_names, XlsxDefinedNames};
use quick_xml::events::Event;
use sheet::{load_sheets, XlsxSheets};
use workbook_properties::XlsxWorkbookProperties;
use workbook_view::{load_bookviews, XlsxWorkbookViews};
use zip::ZipArchive;

use crate::excel::xml_reader;

pub mod calculation_properties;
pub mod custom_workbook_view;
pub mod defined_name;
pub mod sheet;
pub mod workbook_properties;
pub mod workbook_view;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbook?view=openxml-3.0.1
///
/// Root element of the workbook part
///
/// Example
/// ```
/// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
/// <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/5/main" mlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
///     <fileVersion lastEdited="4" lowestEdited="4" rupBuild="4017"/>
///     <workbookPr dateCompatibility="false" vbName="ThisWorkbook" defaultThemeVersion="123820"/>
///     <bookViews>
///         <workbookView xWindow="120" yWindow="45" windowWidth="15135" windowHeight="7650" activeTab="4"/>
///     </bookViews>
///     <sheets>
///         <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
///         <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
///         <sheet name="Sheet5" sheetId="3" r:id="rId3"/>
///         <sheet name="Chart1" sheetId="4" type="chartsheet" r:id="rId4"/>
///     </sheets>
///     <definedNames>
///         <definedName name="MyDefinedName">Sheet3!$A$1:$C$12</definedName>
///     </definedNames>
///     <calcPr calcId="122211" calcMode="autoNoTable" refMode="R1C1" iterate="1" fullPrecision="0"/>
///     <customWorkbookViews>
///         <customWorkbookView name="CustomView1" guid="{CE6681F1-E999-414D-8446-68A031534B57}" maximized="1" xWindow="1"       yWindow="1" windowWidth="1024" windowHeight="547" activeSheetId="1"/>
///     </customWorkbookViews>
///     <pivotCaches>
///         <pivotCache cacheId="0" r:id="rId8"/>
///     </pivotCaches>
///     <smartTagPr embed="1" show="noIndicator"/>
///     <smartTagTypes>
///         <smartTagType namespaceUri="urn:schemas-openxmlformats-org:office:smarttags"       name="date"/>
///     </smartTagTypes>
///     <webPublishing codePage="1252"/>
/// </workbook>
/// ```
/// xml tag: workbook
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxWorkbook {
    // extLst (Future Feature Data Storage Area)	Not supported

    // Child Elements
    // bookViews (Workbook Views)	§18.2.1
    pub bookviews: Option<XlsxWorkbookViews>,

    // calcPr (Calculation Properties)	§18.2.2
    pub calculation_propertis: Option<XlsxCalculationProperties>,

    // customWorkbookViews (Custom Workbook Views)	§18.2.4
    pub custom_workbook_views: Option<XlsxCustomWorkbookViews>,

    // definedNames (Defined Names)	§18.2.6
    pub defined_names: Option<XlsxDefinedNames>,
    // externalReferences (External References)	§18.2.9
    // fileRecoveryPr (File Recovery Properties)	§18.2.11
    // fileSharing (File Sharing)	§18.2.12
    // fileVersion (File Version)	§18.2.13
    // functionGroups (Function Groups)	§18.2.15
    // oleSize (Embedded Object Size)	§18.2.16
    // pivotCaches (PivotCaches)	§18.2.18
    // sheets (Sheets)	§18.2.20
    pub sheets: Option<XlsxSheets>,
    // smartTagPr (Smart Tag Properties)	§18.2.21
    // smartTagTypes (Smart Tag Types)	§18.2.23
    // webPublishing (Web Publishing Properties)	§18.2.24
    // webPublishObjects (Web Publish Objects)	§18.2.26
    // workbookPr (Workbook Properties)	§18.2.28
    pub workbook_properties: Option<XlsxWorkbookProperties>,
    // workbookProtection (Workbook Protection)
}

impl XlsxWorkbook {
    pub(crate) fn load(zip: &mut ZipArchive<impl Read + Seek>) -> anyhow::Result<Self> {
        let path = "xl/workbook.xml";
        let mut workbook = Self {
            bookviews: None,
            calculation_propertis: None,
            custom_workbook_views: None,
            defined_names: None,
            sheets: None,
            workbook_properties: None,
        };

        let Some(mut reader) = xml_reader(zip, path) else {
            return Ok(workbook);
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bookViews" => {
                    workbook.bookviews = Some(load_bookviews(&mut reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"calcPr" => {
                    workbook.calculation_propertis = Some(XlsxCalculationProperties::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"customWorkbookViews" => {
                    workbook.custom_workbook_views = Some(load_custom_bookviews(&mut reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"definedNames" => {
                    workbook.defined_names = Some(load_defined_names(&mut reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheets" => {
                    workbook.sheets = Some(load_sheets(&mut reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"workbookPr" => {
                    workbook.workbook_properties = Some(XlsxWorkbookProperties::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"workbook" => break,
                Ok(Event::Eof) => break,
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(workbook);
    }
}
