use anyhow::{bail, Context};
use quick_xml::events::{BytesStart, Event};
use std::io::{Read, Seek};
use zip::ZipArchive;

use crate::excel::xml_reader;

pub(crate) const EXTERNAL_TARGET_MODE: &str = "External";

/// Example
/// ```
/// <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
///   <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
///   <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
///   <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />
///   <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
/// </Relationships>
/// ```
pub type XlsxRelationships = Vec<XlsxRelationship>;

/// get relationships of a workbook
pub(crate) fn load_workbook_relationships(
    zip: &mut ZipArchive<impl Read + Seek>,
) -> anyhow::Result<XlsxRelationships> {
    let path = "xl/_rels/workbook.xml.rels";
    let Some(mut reader) = xml_reader(zip, path) else {
        bail!("Failed to get relationships.");
    };

    let mut buf = Vec::new();
    let mut relationships: Vec<XlsxRelationship> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"Relationship" => {
                let Some(rel) = XlsxRelationship::load(e)? else {
                    continue;
                };

                relationships.push(rel);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"Relationships" => break,
            Ok(Event::Eof) => break,
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(relationships)
}

/// get relationships of a specific sheet within a workbook
pub(crate) fn load_sheet_relationships(
    zip: &mut ZipArchive<impl Read + Seek>,
    sheet_path: &str,
) -> anyhow::Result<XlsxRelationships> {
    let last_folder_index = sheet_path
        .rfind('/')
        .context("sheet is not within a folder.")?;
    let (base_folder, file_name) = sheet_path.split_at(last_folder_index);
    let path = format!("{}/_rels{}.rels", base_folder, file_name);

    let Some(mut reader) = xml_reader(zip, &path) else {
        bail!("Relationships does not exist for sheet {}.", sheet_path);
    };

    let mut buf = Vec::new();
    let mut relationships: Vec<XlsxRelationship> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"Relationship" => {
                let Some(mut rel) = XlsxRelationship::load(e)? else {
                    continue;
                };
                let target = rel.clone().target;
                // format relative paths
                if target.starts_with("../") {
                    let new_index = base_folder
                        .rfind('/')
                        .context("base folder is not within a parent folder.")?;
                    let full_path = format!("{}{}", &base_folder[..new_index], &target[2..]);
                    rel.target = full_path;
                }

                relationships.push(rel);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"Relationships" => break,
            Ok(Event::Eof) => break,
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(relationships)
}

/// get relationships of a specific drawing within a workbook
pub(crate) fn load_drawing_relationships(
    zip: &mut ZipArchive<impl Read + Seek>,
    drawing_path: &str,
) -> anyhow::Result<XlsxRelationships> {
    let last_folder_index = drawing_path
        .rfind('/')
        .context("drawing is not within a folder.")?;
    let (base_folder, file_name) = drawing_path.split_at(last_folder_index);
    let path = format!("{}/_rels{}.rels", base_folder, file_name);

    let Some(mut reader) = xml_reader(zip, &path) else {
        bail!("Relationships does not exist for drawing {}.", drawing_path);
    };

    let mut buf = Vec::new();
    let mut relationships: Vec<XlsxRelationship> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"Relationship" => {
                let Some(mut rel) = XlsxRelationship::load(e)? else {
                    continue;
                };
                let target = rel.clone().target;
                // format relative paths
                if target.starts_with("../") {
                    let new_index = base_folder
                        .rfind('/')
                        .context("base folder is not within a parent folder.")?;
                    let full_path = format!("{}{}", &base_folder[..new_index], &target[2..]);
                    rel.target = full_path;
                }

                relationships.push(rel);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"Relationships" => break,
            Ok(Event::Eof) => break,
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(relationships)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.ipackagerelationship?view=openxml-3.0.1
///
/// defines an association between a source Package or PackagePart to a target PackagePart or external resource.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxRelationship {
    pub id: String,
    pub r#type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

impl XlsxRelationship {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Option<Self>> {
        let attributes = e.attributes();

        let mut id: Option<String> = None;
        let mut r#type: Option<String> = None;
        let mut target: Option<String> = None;
        let mut target_mode: Option<String> = None;

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"Id" => id = Some(string_value),
                        b"Type" => r#type = Some(string_value),
                        b"Target" => target = Some(string_value),
                        b"TargetMode" => target_mode = Some(string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        if id.is_none() || r#type.is_none() || target.is_none() {
            return Ok(None);
        }

        Ok(Some(Self {
            id: id.unwrap(),
            r#type: r#type.unwrap(),
            target: target.unwrap(),
            target_mode,
        }))
    }
}

/// (id, formated path)
pub(crate) fn zip_path_for_type(
    relationships: &Vec<XlsxRelationship>,
    r#type: &str,
) -> Vec<(String, String)> {
    let filtered: Vec<(String, String)> = relationships
        .iter()
        .filter(|r| {
            r.to_owned()
                .r#type
                .to_lowercase()
                .contains(&r#type.to_lowercase())
        })
        .map(|r| (r.id.to_string(), format_target_path(&r.target)))
        .collect();
    return filtered;
}

pub(crate) fn zip_path_for_id(relationships: &Vec<XlsxRelationship>, id: &str) -> Option<String> {
    let Some(rel) = rel_for_id(relationships, id) else {
        return None;
    };
    return Some(format_target_path(&rel.target));
}

pub(crate) fn raw_target_for_id(relationships: &Vec<XlsxRelationship>, id: &str) -> Option<String> {
    let Some(rel) = rel_for_id(relationships, id) else {
        return None;
    };
    return Some(rel.target);
}

pub(crate) fn rel_for_id(
    relationships: &Vec<XlsxRelationship>,
    id: &str,
) -> Option<XlsxRelationship> {
    let filtered: Vec<XlsxRelationship> = relationships
        .to_owned()
        .into_iter()
        .filter(|r| r.to_owned().id.eq_ignore_ascii_case(&id))
        .collect();
    return filtered.first().cloned();
}

pub(crate) fn format_target_path(target: &str) -> String {
    return if target.starts_with("/xl/") {
        target[1..].to_string()
    } else if target.starts_with("xl/") {
        target.to_string()
    } else if target.starts_with("/") {
        format!("xl{}", target)
    } else {
        format!("xl/{}", target)
    };
}
