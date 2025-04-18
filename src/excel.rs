use anyhow::{bail, Context};
use quick_xml::Reader;
use std::{
    collections::BTreeMap,
    fs::File,
    io::{BufReader, Read, Seek},
    path::Path,
};
use zip::{read::ZipFile, ZipArchive};

use crate::{
    packaging::relationship::{
        load_sheet_relationships, load_workbook_relationships, raw_target_for_id, zip_path_for_id,
        zip_path_for_type, XlsxRelationships,
    },
    processed::spreadsheet::{
        sheet::worksheet::{calculation_reference::CalculationReferenceMode, Worksheet},
        sheet_basic_info::{SheetBasicInfo, SheetType},
    },
    raw::{
        drawing::{scheme::color_scheme::XlsxColorScheme, theme::XlsxTheme},
        spreadsheet::{
            shared_string::shared_string_table::XlsxSharedStringTable,
            sheet::worksheet::XlsxWorksheet, stylesheet::XlsxStyleSheet, table::XlsxTable,
            workbook::XlsxWorkbook,
        },
    },
};

pub(crate) type XmlReader<'a> = Reader<BufReader<ZipFile<'a>>>;

/// A struct representing xml zipped excel file
pub struct Excel<RS> {
    zip: ZipArchive<RS>,
    workbook_relationships: XlsxRelationships,
    stylesheet: Box<Option<XlsxStyleSheet>>,
    theme: Box<Option<XlsxTheme>>,
    shared_strings: Box<Option<XlsxSharedStringTable>>,
    workbook: Box<Option<XlsxWorkbook>>,
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
            stylesheet: Box::new(None),
            theme: Box::new(None),
            shared_strings: Box::new(None),
            workbook: Box::new(None),
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
    pub fn get_raw_stylesheet(&mut self) -> anyhow::Result<Box<Option<XlsxStyleSheet>>> {
        if self.stylesheet.is_none() {
            self.stylesheet = Box::new(Some(XlsxStyleSheet::load(&mut self.zip)?));
        }
        return Ok(self.stylesheet.clone());
    }

    /// Get theme used.
    /// Parsed from get stylesheet parsed from xl/theme/theme{}.xml
    pub fn get_raw_theme(&mut self) -> anyhow::Result<Box<Option<XlsxTheme>>> {
        if self.theme.is_none() {
            let path = zip_path_for_type(&self.workbook_relationships, "theme");
            self.theme = Box::new(Some(XlsxTheme::load(&mut self.zip, path)?));
        }
        return Ok(self.theme.clone());
    }

    /// Get shared string parsed from xl/sharedStrings.xml
    pub fn get_raw_shared_strings(&mut self) -> anyhow::Result<Box<Option<XlsxSharedStringTable>>> {
        if self.shared_strings.is_none() {
            self.shared_strings = Box::new(Some(XlsxSharedStringTable::load(&mut self.zip)?));
        }
        return Ok(self.shared_strings.clone());
    }

    /// Get workbook parsed from xl/workbook.xml
    pub fn get_raw_workbook(&mut self) -> anyhow::Result<Box<Option<XlsxWorkbook>>> {
        if self.workbook.is_none() {
            self.workbook = Box::new(Some(XlsxWorkbook::load(&mut self.zip)?));
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
}

/// functions for getting processed parsed results
impl<RS: Read + Seek> Excel<RS> {
    /// Get a list of sheets in the workbook
    pub fn get_sheets(&mut self) -> anyhow::Result<Vec<SheetBasicInfo>> {
        let Some(workbook) = *self.get_raw_workbook()?.clone() else {
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

        let shared_strings = if let Some(table) = *self.get_raw_shared_strings()? {
            table.string_item.unwrap_or(vec![])
        } else {
            vec![]
        };

        let stylesheet = self
            .get_raw_stylesheet()?
            .context("Style sheet not availalble")?;

        let mut color_scheme: Option<XlsxColorScheme> = None;
        if let Some(theme) = *self.get_raw_theme()? {
            if let Some(theme_elements) = theme.theme_elements {
                color_scheme = theme_elements.color_scheme
            }
        };

        let tables = self.get_raw_tables(raw_worksheet.clone(), worksheet_rels.clone())?;
        let rel_hyperlinks =
            self.get_hyperlinks_in_rel(raw_worksheet.clone(), worksheet_rels.clone())?;

        let worksheet = Worksheet::from_raw(
            sheet.clone().name,
            sheet.sheet_id,
            raw_worksheet,
            tables,
            raw_workbook.clone().defined_names.unwrap_or(vec![]),
            rel_hyperlinks,
            self.is_1904(raw_workbook.clone()),
            self.calculation_mode(raw_workbook.clone()),
            shared_strings,
            stylesheet.clone(),
            color_scheme,
        );

        Ok(worksheet)
    }
}

/// private helper functions
impl<RS: Read + Seek> Excel<RS> {
    /// get a list of hyperlinks defined in a worksheet relationships
    ///
    /// (r_id, target): Example: `("rId1", "www.google.com")`
    fn get_hyperlinks_in_rel(
        &self,
        raw_worksheet: XlsxWorksheet,
        worksheet_rels: XlsxRelationships,
    ) -> anyhow::Result<BTreeMap<String, String>> {
        let hyperlinks = raw_worksheet.hyperlinks.unwrap_or(vec![]);
        if hyperlinks.is_empty() {
            return Ok(BTreeMap::new());
        } else {
            let rels: BTreeMap<String, String> = hyperlinks
                .into_iter()
                .filter(|h| h.r_id.is_some())
                .map(|h| {
                    (
                        h.r_id.clone().unwrap(),
                        raw_target_for_id(&worksheet_rels, &h.r_id.clone().unwrap()),
                    )
                })
                .filter(|p| p.1.is_some())
                .map(|p| (p.0, p.1.unwrap()))
                .collect();
            return Ok(rels);
        };
    }

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
) -> Option<XmlReader<'a>> {
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
