use std::collections::BTreeMap;

use regex::Regex;

use crate::{
    common_types::{Coordinate, Dimension},
    raw::spreadsheet::{
        sheet::worksheet::hyperlink::XlsxHyperlink, workbook::defined_name::XlsxDefinedNames,
    },
};

#[derive(Debug, Clone, PartialEq)]
pub enum Hyperlink {
    Inernal(InternalHyperlink),
    External(ExternalHyperlink),
}

impl Hyperlink {
    /// worksheet_rel: (r_id: Target)
    pub(crate) fn from_raw(
        hyperlink: XlsxHyperlink,
        worksheet_rel: BTreeMap<String, String>,
        defined_names: XlsxDefinedNames,
    ) -> Option<Self> {
        if let Some(r_id) = hyperlink.r_id {
            if let Some(v) = worksheet_rel.get(&r_id) {
                return Some(Self::External(ExternalHyperlink::from_string(v)));
            }
        }

        if let Some(location) = hyperlink.location {
            // defined names
            let target_name: XlsxDefinedNames = defined_names
                .into_iter()
                .filter(|n| n.name == Some(location.clone()) && n.value.is_some())
                .collect();

            if let Some(target_name) = target_name.first() {
                return Some(Self::Inernal(InternalHyperlink::from_location_string(
                    &target_name.value.clone().unwrap_or("".to_string()),
                )));
            }

            // direct reference
            return Some(Self::Inernal(InternalHyperlink::from_location_string(
                &location,
            )));
        }
        return None;
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct InternalHyperlink {
    pub sheet_name: String,
    pub cell_range: Dimension,
}

impl InternalHyperlink {
    /// Example location:
    ///
    /// without defined names
    /// * `Sheet1!H7`
    /// * `'Sheet some space 2'!A1`
    /// * `Sheet1!R1C1`
    /// * `'Sheet 2 - Custom grid lines'!R2C1`
    ///
    /// Defined names:
    /// * 'Sheet some space 2'!$B$7:$B$11
    /// * Sheet1!$H$13
    pub(crate) fn from_location_string(location: &str) -> Self {
        // sheet name only
        let location_trimmed = location.trim_start_matches("'").trim_end_matches("'");
        let default = Self {
            sheet_name: location_trimmed.to_string(),
            cell_range: Dimension {
                start: Coordinate { row: 1, col: 1 },
                end: Coordinate { row: 1, col: 1 },
            },
        };

        // without defined name
        let Ok(re) = Regex::new(r"('?)(?<name>.+?)('?)!(?<ref>.*?)$") else {
            return default;
        };

        if let Some(caps) = re.captures(location) {
            let name = caps["name"].to_string();
            let reference = caps["ref"].to_string().replace("$", "");
            // refer to a dimension
            let dimension = if let Some(d) = Dimension::from_r1c1(&reference) {
                Some(d)
            } else if let Some(d) = Dimension::from_a1(reference.as_bytes()) {
                Some(d)
            } else {
                None
            };
            if let Some(dimension) = dimension {
                return Self {
                    sheet_name: name,
                    cell_range: dimension,
                };
            };

            // refer to a single cell
            let coordinate = if let Some(c) = Coordinate::from_r1c1(&reference) {
                c
            } else if let Some(c) = Coordinate::from_a1(reference.as_bytes()) {
                c
            } else {
                Coordinate::from_point((1, 1))
            };

            return Self {
                sheet_name: name,
                cell_range: Dimension {
                    start: coordinate,
                    end: coordinate,
                },
            };
        };

        return default;
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ExternalHyperlink {
    Url(String),
    Email(EmailHyperlink),
}

impl ExternalHyperlink {
    pub(crate) fn from_string(s: &str) -> Self {
        // mail with subject
        if let Ok(re) = Regex::new(r"mailto:(?<email>.+?)(\?subject=)(?<subject>.+?)$") {
            if let Some(caps) = re.captures(s) {
                let email = &caps["email"];
                let email = urlencoding::decode(email)
                    .unwrap_or(std::borrow::Cow::Borrowed(email))
                    .to_string();

                let subject = &caps["subject"];
                let subject = urlencoding::decode(subject)
                    .unwrap_or(std::borrow::Cow::Borrowed(subject))
                    .to_string();

                return Self::Email(EmailHyperlink {
                    mail_to: email,
                    subject,
                });
            };
        };

        // mail without subject
        if let Ok(re) = Regex::new(r"mailto:(?<email>.+?)$") {
            if let Some(caps) = re.captures(s) {
                let email = &caps["email"];
                let email = urlencoding::decode(email)
                    .unwrap_or(std::borrow::Cow::Borrowed(email))
                    .to_string();
                return Self::Email(EmailHyperlink {
                    mail_to: email,
                    subject: "".to_string(),
                });
            };
        };

        let url = urlencoding::decode(s)
            .unwrap_or(std::borrow::Cow::Borrowed(s))
            .to_string();

        return Self::Url(url);
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct EmailHyperlink {
    mail_to: String,
    subject: String,
}
