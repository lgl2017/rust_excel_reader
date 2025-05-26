#[cfg(feature = "drawing")]
use std::collections::BTreeMap;

use anyhow::{bail, Context};
use quick_xml::Reader;
use std::{
    fs::File,
    io::{BufReader, Read, Seek},
    path::Path,
};

use zip::{read::ZipFile, ZipArchive};

#[cfg(feature = "drawing")]
use crate::packaging::relationship::load_drawing_relationships;

#[cfg(feature = "drawing")]
use crate::raw::drawing::worksheet_drawing::XlsxWorksheetDrawing;

use crate::{
    packaging::relationship::{
        load_sheet_relationships, load_workbook_relationships, zip_path_for_id, zip_path_for_type,
        XlsxRelationships,
    },
    processed::spreadsheet::{
        sheet::worksheet::{calculation_reference::CalculationReferenceMode, Worksheet},
        sheet_basic_info::{SheetBasicInfo, SheetType},
    },
    raw::{
        drawing::theme::XlsxTheme,
        spreadsheet::{
            shared_string::shared_string_table::XlsxSharedStringTable,
            sheet::worksheet::XlsxWorksheet, stylesheet::XlsxStyleSheet, table::XlsxTable,
            workbook::XlsxWorkbook,
        },
    },
};

pub(crate) type XmlReader<'a, R> = Reader<BufReader<ZipFile<'a, R>>>;

/// A struct representing xml zipped excel file
pub struct Excel<RS> {
    zip: ZipArchive<RS>,
    workbook_relationships: XlsxRelationships,
    stylesheet: Option<Box<XlsxStyleSheet>>,
    theme: Option<Box<XlsxTheme>>,
    shared_strings: Option<Box<XlsxSharedStringTable>>,
    workbook: Option<Box<XlsxWorkbook>>,
}

// initialization
impl Excel<BufReader<File>> {
    pub fn from_path<P: AsRef<Path>>(path: P) -> anyhow::Result<Excel<BufReader<File>>> {
        let reader = BufReader::new(File::open(path)?);
        return Self::from_reader(reader);
    }
}

impl<RS: Read + Seek> Excel<RS> {
    pub fn from_reader(reader: RS) -> anyhow::Result<Excel<RS>> {
        let mut zip = ZipArchive::new(reader)?;
        let relationships = load_workbook_relationships(&mut zip)?;
        Ok(Self {
            zip,
            workbook_relationships: relationships,
            stylesheet: None,
            theme: None,
            shared_strings: None,
            workbook: None,
        })
    }
}

/// functions for getting raw parsed results
impl<RS: Read + Seek> Excel<RS> {
    /// Get relationship parsed from xl/_rels/workbook.xml.rels
    pub fn get_raw_workbook_relationship(&mut self) -> XlsxRelationships {
        return self.workbook_relationships.clone();
    }

    /// Get stylesheet parsed from xl/styles.xml
    pub fn get_raw_stylesheet(&mut self) -> anyhow::Result<Option<Box<XlsxStyleSheet>>> {
        if self.stylesheet.is_none() {
            self.stylesheet = Some(Box::new(XlsxStyleSheet::load(&mut self.zip)?));
        }
        return Ok(self.stylesheet.clone());
    }

    /// Get theme used.
    /// Parsed from get stylesheet parsed from xl/theme/theme{}.xml
    pub fn get_raw_theme(&mut self) -> anyhow::Result<Option<Box<XlsxTheme>>> {
        if self.theme.is_none() {
            let path = zip_path_for_type(&self.workbook_relationships, "theme");
            let path = path.iter().map(|p| p.1.to_string()).collect();
            self.theme = Some(Box::new(XlsxTheme::load(&mut self.zip, path)?));
        }
        return Ok(self.theme.clone());
    }

    /// Get shared string parsed from xl/sharedStrings.xml
    pub fn get_raw_shared_strings(&mut self) -> anyhow::Result<Option<Box<XlsxSharedStringTable>>> {
        if self.shared_strings.is_none() {
            self.shared_strings = Some(Box::new(XlsxSharedStringTable::load(&mut self.zip)?));
        }
        return Ok(self.shared_strings.clone());
    }

    /// Get workbook parsed from xl/workbook.xml
    pub fn get_raw_workbook(&mut self) -> anyhow::Result<Option<Box<XlsxWorkbook>>> {
        if self.workbook.is_none() {
            self.workbook = Some(Box::new(XlsxWorkbook::load(&mut self.zip)?));
        }
        return Ok(self.workbook.clone());
    }

    /// Get a specific worksheet parsed from xl/worksheets/sheet{}.xml
    ///
    /// * name: worksheet name
    pub fn get_raw_worksheet_with_name(&mut self, name: &str) -> anyhow::Result<XlsxWorksheet> {
        let sheet = self.get_sheet_with_name(name)?;
        return self.get_raw_worksheet(&sheet);
    }

    /// Get a specific worksheet parsed from xl/worksheets/sheet{}.xml
    ///
    /// * id: worksheet sheet id
    pub fn get_raw_worksheet_with_sheet_id(&mut self, id: &u64) -> anyhow::Result<XlsxWorksheet> {
        let sheet = self.get_sheet_with_sheet_id(id)?;
        return self.get_raw_worksheet(&sheet);
    }

    /// Get a specific worksheet parsed from xl/worksheets/sheet{}.xml
    pub fn get_raw_worksheet(&mut self, sheet: &SheetBasicInfo) -> anyhow::Result<XlsxWorksheet> {
        if sheet.r#type != SheetType::WorkSheet {
            bail!("Sheet specified is not a worksheet")
        };
        return XlsxWorksheet::load(&mut self.zip, &sheet.path);
    }

    /// Get relationships for a sheet parsed from xl/worksheets/_rels/sheet{}.xml.rels
    ///
    /// * name: worksheet name
    pub fn get_raw_sheet_relationship_with_name(
        &mut self,
        name: &str,
    ) -> anyhow::Result<XlsxRelationships> {
        let sheet = self.get_sheet_with_name(name)?;
        return self.get_raw_sheet_relationship(&sheet);
    }

    /// Get relationships for a sheet parsed from xl/worksheets/_rels/sheet{}.xml.rels
    ///
    /// * id: worksheet sheet id
    pub fn get_raw_sheet_relationship_with_sheet_id(
        &mut self,
        id: &u64,
    ) -> anyhow::Result<XlsxRelationships> {
        let sheet = self.get_sheet_with_sheet_id(id)?;
        return self.get_raw_sheet_relationship(&sheet);
    }

    /// Get relationship for a sheet parsed from xl/worksheets/_rels/sheet{}.xml.rels
    pub fn get_raw_sheet_relationship(
        &mut self,
        sheet: &SheetBasicInfo,
    ) -> anyhow::Result<XlsxRelationships> {
        let worksheet_rels = load_sheet_relationships(&mut self.zip, &sheet.path)?;
        return Ok(worksheet_rels);
    }

    /// Get a specific worksheet parsed from xl/worksheets/sheet{}.xml
    ///
    /// * name: worksheet name
    pub fn get_raw_tables_for_worksheet_with_name(
        &mut self,
        name: &str,
    ) -> anyhow::Result<Vec<XlsxTable>> {
        let sheet = self.get_sheet_with_name(name)?;
        return self.get_raw_tables_for_worksheet(&sheet);
    }

    /// Get tables defined in a worksheet parsed from xl/tables/table{}.xml, ..., xl/tables/table{n}.xml,
    ///
    /// * id: worksheet sheet id
    pub fn get_raw_tables_for_worksheet_with_sheet_id(
        &mut self,
        id: &u64,
    ) -> anyhow::Result<Vec<XlsxTable>> {
        let sheet = self.get_sheet_with_sheet_id(id)?;
        return self.get_raw_tables_for_worksheet(&sheet);
    }

    /// Get tables defined in a worksheet parsed from xl/tables/table{}.xml, ..., xl/tables/table{n}.xml,
    pub fn get_raw_tables_for_worksheet(
        &mut self,
        sheet: &SheetBasicInfo,
    ) -> anyhow::Result<Vec<XlsxTable>> {
        let raw_worksheet = self.get_raw_worksheet(&sheet)?;
        let worksheet_rels = self.get_raw_sheet_relationship(&sheet).unwrap_or(vec![]);
        return self.get_raw_tables(raw_worksheet, worksheet_rels);
    }

    /// Get XlsxWorksheetDrawing that defines all drawing objects within the worksheet parsed from xl/drawings/drawing{}.xml
    #[cfg(feature = "drawing")]
    pub fn get_raw_drawing_for_worksheet(
        &mut self,
        sheet: &SheetBasicInfo,
    ) -> anyhow::Result<Option<(XlsxWorksheetDrawing, XlsxRelationships)>> {
        let raw_worksheet = self.get_raw_worksheet(&sheet)?;
        let worksheet_rels = self.get_raw_sheet_relationship(&sheet).unwrap_or(vec![]);
        return self.get_raw_drawing(raw_worksheet, worksheet_rels);
    }
}

/// functions for getting processed parsed results
impl<RS: Read + Seek> Excel<RS> {
    /// Get a list of sheets in the workbook
    pub fn get_sheets(&mut self) -> anyhow::Result<Vec<SheetBasicInfo>> {
        let Some(workbook) = self.get_raw_workbook()?.clone() else {
            return Ok(vec![]);
        };
        let Some(sheets) = workbook.sheets.clone() else {
            return Ok(vec![]);
        };
        let sheets: anyhow::Result<Vec<SheetBasicInfo>> = sheets
            .iter()
            .map(|s| SheetBasicInfo::from_raw(s.clone(), &self.workbook_relationships))
            .collect();

        return sheets;
    }

    /// Get worksheet (processed)
    ///
    /// name: Worksheet name
    pub fn get_worksheet_with_name(&mut self, name: &str) -> anyhow::Result<Worksheet> {
        let sheet = self.get_sheet_with_name(name)?;
        return self.get_worksheet(&sheet);
    }

    /// Get worksheet (processed)
    ///
    /// id: Worksheet sheet id
    pub fn get_worksheet_with_sheet_id(&mut self, id: &u64) -> anyhow::Result<Worksheet> {
        let sheet = self.get_sheet_with_sheet_id(id)?;
        return self.get_worksheet(&sheet);
    }

    /// Get worksheet (processed)
    pub fn get_worksheet(&mut self, sheet: &SheetBasicInfo) -> anyhow::Result<Worksheet> {
        let raw_workbook = self.get_raw_workbook()?.context("workbook not available")?;
        let raw_worksheet = self.get_raw_worksheet(&sheet)?;
        let worksheet_rels = self.get_raw_sheet_relationship(&sheet).unwrap_or(vec![]);

        let shared_strings = if let Some(table) = self.get_raw_shared_strings()? {
            table.string_item.unwrap_or(vec![])
        } else {
            vec![]
        };

        let stylesheet = self
            .get_raw_stylesheet()?
            .context("Style sheet not availalble")?;

        let theme = self.get_raw_theme()?;

        let tables = self.get_raw_tables(raw_worksheet.clone(), worksheet_rels.clone())?;

        #[cfg(feature = "drawing")]
        let mut drawing_rel: XlsxRelationships = vec![];
        #[cfg(feature = "drawing")]
        let mut raw_drawing: Option<Box<XlsxWorksheetDrawing>> = None;

        #[cfg(feature = "drawing")]
        if let Some(drawing) =
            self.get_raw_drawing(raw_worksheet.clone(), worksheet_rels.clone())?
        {
            drawing_rel = drawing.1;
            raw_drawing = Some(Box::new(drawing.0));
        }
        #[cfg(feature = "drawing")]
        let bytes = self.get_image_bytes_in_rel(drawing_rel.clone());

        let worksheet = Worksheet::from_raw(
            sheet.clone().name,
            sheet.sheet_id,
            Box::new(raw_worksheet),
            Box::new(worksheet_rels),
            Box::new(tables),
            Box::new(raw_workbook.clone().defined_names.unwrap_or(vec![])),
            self.is_1904(*raw_workbook.clone()),
            self.calculation_mode(*raw_workbook.clone()),
            Box::new(shared_strings),
            stylesheet.clone(),
            theme.clone(),
            #[cfg(feature = "drawing")]
            Box::new(drawing_rel),
            #[cfg(feature = "drawing")]
            raw_drawing,
            #[cfg(feature = "drawing")]
            Box::new(bytes),
        );

        Ok(worksheet)
    }
}

/// private helper functions
impl<RS: Read + Seek> Excel<RS> {
    /// get a list of tables used in a worksheet
    fn get_raw_tables(
        &mut self,
        raw_worksheet: XlsxWorksheet,
        worksheet_rels: XlsxRelationships,
    ) -> anyhow::Result<Vec<XlsxTable>> {
        let table_parts = raw_worksheet.table_parts.unwrap_or(vec![]);
        if table_parts.is_empty() {
            return Ok(vec![]);
        } else {
            let paths: Vec<String> = table_parts
                .into_iter()
                .map(|t| zip_path_for_id(&worksheet_rels, &t.id))
                .filter(|p| p.is_some())
                .map(|p| p.unwrap())
                .collect();

            let raw_tables: Vec<XlsxTable> = paths
                .into_iter()
                .map(|p| XlsxTable::load(&mut self.zip, &p))
                .filter(|t| t.is_ok())
                .map(|t| t.unwrap())
                .collect();

            return Ok(raw_tables);
        };
    }

    /// get
    /// - `XlsxWorksheetDrawing` parsed from xl/drawings/drawing{}.xml that defines all drawing objects within the worksheet
    /// - `Relationship` from the xl/drawings/_rels/drawing{}.xml.rel
    #[cfg(feature = "drawing")]
    fn get_raw_drawing(
        &mut self,
        raw_worksheet: XlsxWorksheet,
        worksheet_rels: XlsxRelationships,
    ) -> anyhow::Result<Option<(XlsxWorksheetDrawing, XlsxRelationships)>> {
        let Some(drawing) = raw_worksheet.drawing else {
            return Ok(None);
        };
        let Some(path) = zip_path_for_id(&worksheet_rels, &drawing.id) else {
            return Ok(None);
        };
        let drawing_rels = load_drawing_relationships(&mut self.zip, &path).unwrap_or(vec![]);
        return Ok(Some((
            XlsxWorksheetDrawing::load(&mut self.zip, &path)?,
            drawing_rels,
        )));
    }

    /// get a list of image bytes defined in a drawing relationships
    ///
    /// (r_id, bytes): Example: `("rId1", some bytes)`
    #[cfg(feature = "drawing")]
    fn get_image_bytes_in_rel(
        &mut self,
        drawing_rel: XlsxRelationships,
    ) -> BTreeMap<String, Vec<u8>> {
        let rels = zip_path_for_type(&drawing_rel, "image");
        let mut bytes: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        for rel in rels.into_iter() {
            if let Ok(b) = self.get_bytes_for_path(&rel.1) {
                bytes.insert(rel.0, b);
            }
        }
        return bytes;
    }

    #[cfg(feature = "drawing")]
    fn get_bytes_for_path(&mut self, path: &str) -> anyhow::Result<Vec<u8>> {
        let zip = &mut self.zip;
        let path = get_actual_path(zip, path)
            .context(format!("File does not exist for path: {}", path))?;
        let mut zip = zip.by_name(&path)?;
        let mut buf: Vec<u8> = Vec::new();
        zip.read_to_end(&mut buf)?;
        Ok(buf)
    }

    fn get_sheet_with_name(&mut self, name: &str) -> anyhow::Result<SheetBasicInfo> {
        let sheets = self.get_sheets()?;
        let target: Vec<SheetBasicInfo> = sheets
            .into_iter()
            .filter(|s| s.name.eq_ignore_ascii_case(name))
            .collect();
        let Some(target) = target.first() else {
            bail!("Sheet with name: `{}` does not exist.", name)
        };
        return Ok(target.to_owned());
    }

    fn get_sheet_with_sheet_id(&mut self, id: &u64) -> anyhow::Result<SheetBasicInfo> {
        let sheets = self.get_sheets()?;
        let target: Vec<SheetBasicInfo> =
            sheets.into_iter().filter(|s| s.sheet_id.eq(id)).collect();
        let Some(target) = target.first() else {
            bail!("Worksheet with id: `{}` does not exist.", id)
        };
        return Ok(target.to_owned());
    }

    fn is_1904(&self, workbook: XlsxWorkbook) -> bool {
        let Some(properties) = workbook.workbook_properties else {
            return false;
        };
        let compatibility = properties.date_compatibility.unwrap_or(true);
        if compatibility == false {
            return false;
        }
        return properties.date1904.unwrap_or(false);
    }

    fn calculation_mode(&self, workbook: XlsxWorkbook) -> Option<CalculationReferenceMode> {
        let Some(properties) = workbook.calculation_propertis else {
            return None;
        };
        return CalculationReferenceMode::from_string(properties.reference_mode);
    }
}

pub(crate) fn xml_reader<'a, RS: Read + Seek>(
    zip: &'a mut ZipArchive<RS>,
    path: &str,
) -> Option<XmlReader<'a, RS>> {
    let Some(path) = get_actual_path(zip, path) else {
        return None;
    };
    let Ok(zip) = zip.by_name(&path) else {
        return None;
    };
    let mut xml_reader = Reader::from_reader(BufReader::new(zip));

    let config = xml_reader.config_mut();
    config.allow_unmatched_ends = false; // default false
    config.check_comments = false; // default false
    config.check_end_names = false; // default true
    config.trim_text(false); // default false
    config.expand_empty_elements = true; // default false

    return Some(xml_reader);
}

fn get_actual_path<'a, RS: Read + Seek>(zip: &'a mut ZipArchive<RS>, path: &str) -> Option<String> {
    return zip
        .file_names()
        .find(|n| n.eq_ignore_ascii_case(path))
        .to_owned()
        .map(|f| f.to_owned());
}
