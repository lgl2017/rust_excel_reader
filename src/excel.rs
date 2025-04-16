use anyhow::{bail, Context};
use quick_xml::Reader;
use std::{
    fs::File,
    io::{BufReader, Read, Seek},
    path::Path,
};
use zip::{read::ZipFile, ZipArchive};

use crate::{
    packaging::relationship::{
        load_workbook_relationships, load_worksheet_relationships, zip_path_for_id,
        zip_path_for_type, Relationships,
    },
    processed::spreadsheet::{
        sheet::worksheet::{calculation_reference::CalculationReferenceMode, Worksheet},
        sheet_basic_info::SheetBasicInfo,
    },
    raw::{
        drawing::{scheme::color_scheme::ColorScheme, theme::Theme},
        spreadsheet::{
            shared_string::shared_string_table::SharedStringTable,
            sheet::worksheet::Worksheet as RawWorksheet, stylesheet::StyleSheet,
            table::Table as RawTable, workbook::Workbook,
        },
    },
};

pub(crate) type XmlReader<'a> = Reader<BufReader<ZipFile<'a>>>;

/// A struct representing xml zipped excel file
pub struct Excel<RS> {
    zip: ZipArchive<RS>,
    workbook_relationships: Relationships,
    stylesheet: Option<Box<StyleSheet>>,
    theme: Option<Box<Theme>>,
    shared_strings: Option<Box<SharedStringTable>>,
    workbook: Option<Box<Workbook>>,
}

// initialization
impl Excel<BufReader<File>> {
    pub fn from_path<P: AsRef<Path>>(path: P) -> anyhow::Result<Excel<BufReader<File>>> {
        let reader = BufReader::new(File::open(path)?);
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
    /// Get stylesheet parsed from xl/styles.xml
    pub fn get_raw_stylesheet(&mut self) -> anyhow::Result<Option<Box<StyleSheet>>> {
        if self.stylesheet.is_none() {
            self.stylesheet = Some(Box::new(StyleSheet::load(&mut self.zip)?));
        }
        return Ok(self.stylesheet.clone());
    }

    /// Get theme used.
    /// Parsed from get stylesheet parsed from xl/theme/theme{}.xml
    pub fn get_raw_theme(&mut self) -> anyhow::Result<Option<Box<Theme>>> {
        if self.theme.is_none() {
            let path = zip_path_for_type(&self.workbook_relationships, "theme");
            self.theme = Some(Box::new(Theme::load(&mut self.zip, path)?));
        }
        return Ok(self.theme.clone());
    }

    /// Get shared string parsed from xl/sharedStrings.xml
    pub fn get_raw_shared_strings(&mut self) -> anyhow::Result<Option<Box<SharedStringTable>>> {
        if self.shared_strings.is_none() {
            self.shared_strings = Some(Box::new(SharedStringTable::load(&mut self.zip)?));
        }
        return Ok(self.shared_strings.clone());
    }

    /// Get workbook parsed from xl/workbook.xml
    pub fn get_raw_workbook(&mut self) -> anyhow::Result<Option<Box<Workbook>>> {
        if self.workbook.is_none() {
            self.workbook = Some(Box::new(Workbook::load(&mut self.zip)?));
        }
        return Ok(self.workbook.clone());
    }

    /// Get a specific worksheet parsed from xl/worksheets/sheet{}.xml
    ///
    /// * name: worksheet name
    pub fn get_raw_worksheet_with_name(&mut self, name: &str) -> anyhow::Result<RawWorksheet> {
        let sheet = self.get_sheet_with_name(name)?;
        return self.get_raw_worksheet(&sheet);
    }

    /// Get a specific worksheet parsed from xl/worksheets/sheet{}.xml
    ///
    /// * id: worksheet sheet id
    pub fn get_raw_worksheet_with_sheet_id(&mut self, id: &u64) -> anyhow::Result<RawWorksheet> {
        let sheet = self.get_sheet_with_sheet_id(id)?;
        return self.get_raw_worksheet(&sheet);
    }

    /// Get a specific worksheet parsed from xl/worksheets/sheet{}.xml
    pub fn get_raw_worksheet(&mut self, sheet: &SheetBasicInfo) -> anyhow::Result<RawWorksheet> {
        return RawWorksheet::load(&mut self.zip, &sheet.path);
    }

    /// Get tables defined in a worksheet parsed from xl/tables/table{}.xml, ..., xl/tables/table{n}.xml,
    pub fn get_raw_tables_for_worksheet(
        &mut self,
        sheet: &SheetBasicInfo,
    ) -> anyhow::Result<Vec<RawTable>> {
        let raw_worksheet = self.get_raw_worksheet(&sheet)?;
        let table_parts = raw_worksheet.table_parts.unwrap_or(vec![]);
        if table_parts.is_empty() {
            return Ok(vec![]);
        } else {
            let worksheet_rels = load_worksheet_relationships(&mut self.zip, &sheet.path)?;

            let paths: Vec<String> = table_parts
                .into_iter()
                .map(|t| zip_path_for_id(&worksheet_rels, &t.id))
                .filter(|p| p.is_some())
                .map(|p| p.unwrap())
                .collect();

            let raw_tables: Vec<RawTable> = paths
                .into_iter()
                .map(|p| RawTable::load(&mut self.zip, &p))
                .filter(|t| t.is_ok())
                .map(|t| t.unwrap())
                .collect();

            return Ok(raw_tables);
        };
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
        let raw_worksheet = self.get_raw_worksheet(sheet)?;
        let shared_strings = if let Ok(Some(table)) = self.get_raw_shared_strings() {
            table.string_item.unwrap_or(vec![])
        } else {
            vec![]
        };
        let stylesheet = self
            .get_raw_stylesheet()?
            .context("Style sheet not availalble")?;
        let mut color_scheme: Option<ColorScheme> = None;
        if let Some(theme) = self.get_raw_theme()? {
            if let Some(theme_elements) = theme.theme_elements {
                color_scheme = theme_elements.color_scheme
            }
        };
        let tables = self.get_raw_tables_for_worksheet(sheet)?;

        let worksheet = Worksheet::from_raw(
            sheet.clone().name,
            sheet.sheet_id,
            raw_worksheet,
            tables,
            self.is_1904(),
            self.calculation_mode(),
            shared_strings,
            *stylesheet.clone(),
            color_scheme,
        );

        Ok(worksheet)
    }
}

/// private helper functions
impl<RS: Read + Seek> Excel<RS> {
    fn get_sheet_with_name(&mut self, name: &str) -> anyhow::Result<SheetBasicInfo> {
        let sheets = self.get_sheets()?;
        let target: Vec<SheetBasicInfo> = sheets
            .into_iter()
            .filter(|s| s.name.eq_ignore_ascii_case(name))
            .collect();
        let Some(target) = target.first() else {
            bail!("Worksheet with name: `{}` does not exist.", name)
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

    fn is_1904(&mut self) -> bool {
        let Ok(Some(workbook)) = self.get_raw_workbook() else {
            return false;
        };
        let Some(properties) = workbook.workbook_properties else {
            return false;
        };
        let compatibility = properties.date_compatibility.unwrap_or(true);
        if compatibility == false {
            return false;
        }
        return properties.date1904.unwrap_or(false);
    }

    fn calculation_mode(&mut self) -> Option<CalculationReferenceMode> {
        let Ok(Some(workbook)) = self.get_raw_workbook() else {
            return None;
        };
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
