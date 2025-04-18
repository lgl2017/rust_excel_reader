use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::{string_to_bool, string_to_int};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties?view=openxml-3.0.1
///
/// This element defines a collection of workbook properties
///
///
/// Example
/// ```
/// <workbookPr dateCompatibility="false" showObjects="none" saveExternalLinkValues="0"  defaultThemeVersion="123820"/>
/// ```
/// workbookPr (Workbook Properties)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxWorkbookProperties {
    //  Attributes	Description
    /// allowRefreshQuery (Allow Refresh Query)
    ///
    /// Specifies a boolean value that indicates whether the application will refresh query tables in this workbook.
    /// A value of 1 or true indicates the application will refresh query tables when the workbook is loaded.
    /// A value of 0 or false indicates the application will not refresh query tables.
    /// The default value for this attribute is false.
    pub allow_refresh_query: Option<bool>,

    /// autoCompressPictures (Auto Compress Pictures)
    ///
    /// Specifies a boolean value that indicates the application automatically compressed pictures in the workbook.
    /// A value of 1 or true indicates the application automatically compresses pictures of the workbook.
    /// When a picture is compressed, the application:
    /// - Reduces resolution (to 96 dots per inch (dpi) for Web and 200 dpi for print), and unnecessary information is discarded.
    /// - Discards extra information. [Example: When a picture has been cropped or resized, the "hidden" parts of the picture are stored in the file. end example]
    /// - Compress the picture, if possible.
    ///
    /// A value of 0 or false indicates the application does not compress pictures in this workbook.
    ///
    /// The default value for this attribute is true.
    pub auto_compress_pictures: Option<bool>,

    /// backupFile (Create Backup File)
    ///
    /// Specifies a boolean value that indicates whether the application creates a backup of the workbook on save.
    /// A value of 1 or true indicates the application creates a backup of the workbook on save.
    /// A value of 0 or false indicates the application does not create a backup.
    /// The default value for this attribute is false.
    pub backup_file: Option<bool>,

    /// checkCompatibility (Check Compatibility On Save)
    ///
    /// Specifies a boolean value that indicates whether the application checks for compatibility when saving this workbook to older file formats.
    /// A value of 1 or true indicates the application performs a compatibility check when saving to legacy binary formats.
    /// A value of 0 or false indicates the application does not perform a compatibility check when saving to legacy binary formats.
    /// The default value for this attribute is false.
    pub check_compatibility: Option<bool>,

    /// codeName (Code Name)
    ///
    /// Specifies the codename of the application that created this workbook. Use this attribute to track file content in incremental releases of the application.
    pub code_name: Option<String>,

    /// date1904 (Date 1904)
    ///
    /// Value that indicates whether to use a 1900 or 1904 date base when converting serial values in the workbook to dates. [Note: If the dateCompatibility attribute is 0 or false, this attribute is ignored. end note]
    /// A value of 1 or true indicates the workbook uses the 1904 backward compatibility date system.
    /// A value of 0 or false indicates the workbook uses a date system based in 1900, as specified by the value of the dateCompatibility attribute.
    /// The default value for this attribute is false.
    pub date1904: Option<bool>,

    /// dateCompatibility (Date Compatibility)
    /// Specifies whether the date base should be treated as a compatibility date base or should support the full 8601 date range.
    /// A value of 1 or true indicates that the date system in use is either the 1900 backward compatibility date base or the 1904 backward compatibility date base, as specified by the value of the date1904 attribute.
    /// A value of 0 or false indicates that the date system is the 1900 date base, based on the 8601 date range.
    /// The default value for this attribute is true.
    pub date_compatibility: Option<bool>,

    /// defaultThemeVersion (Default Theme Version)
    ///
    /// Specifies the default version of themes to apply in the workbook.
    /// The value for defaultThemeVersion depends on the application.
    /// SpreadsheetML defaults to the form [version][build], where [version] refers to the version of the application, and [build] refers to the build of the application when the themes in the user interface changed.
    pub default_theme_version: Option<i64>,

    /// filterPrivacy (Filter Privacy)
    ///
    /// Specifies a boolean value that indicates whether the application has inspected the workbook for personally identifying information ().
    /// If this flag is set, the application warns the user any time the user performs an action that will insert into the document.
    /// A value of 1 or true indicates the application will warn the user when they insert into the workbook.
    /// A value of 0 or false indicates the application will not warn the user when they insert into the workbook; the workbook has not been inspected for .
    /// The default value for this attribute is false.
    pub filter_privacy: Option<bool>,

    /// hidePivotFieldList (Hide Pivot Field List)
    ///
    /// Specifies a boolean value that indicates whether a list of fields is shown for pivot tables in the application user interface.
    /// A value of 1 or true indicates a list of fields is show for pivot tables.
    /// A value of 0 or false indicates a list of fields is not shown for pivot tables.
    /// The default value for this attribute is false.
    pub hide_pivot_field_list: Option<bool>,

    /// promptedSolutions (Prompted Solutions)
    ///
    /// Specifies a boolean value that indicates whether the user has received an alert to load Smart Document components.
    /// A value of 1 or true indicates the user received an alert to load SmartDoc.
    /// A value of 0 or false indicates the user did not receive an alert.
    /// The default value for this attribute is false.
    pub prompted_solutions: Option<bool>,

    /// publishItems (Publish Items)
    ///
    /// Specifies a boolean value that indicates whether the publish the workbook or workbook items to the application server.
    /// A value of 1 or true indicates that workbook items are published.
    /// A value of 0 or false indicates that the workbook is published.
    /// The default value for this attribute is false.
    pub publish_items: Option<bool>,

    /// refreshAllConnections (Refresh all Connections on Open)
    ///
    /// Specifies a boolean value that indicates whether the workbok shall refresh all the connections to data sources during load.
    /// The default value for this attribute is false.
    pub refresh_all_connections: Option<bool>,

    /// saveExternalLinkValues (Save External Link Values)
    ///
    /// Specifies a boolean value that indicates whether the application will cache values retrieved from other workbooks via an externally linking formula. Data is cached at save.
    /// A value of 1 or true indicates data from externally linked formulas is cached. A supporting part is written out containing a cached cell table from the external workbook.
    /// A value of 0 or false indicates data from externally linked formulas is not cached.
    /// The default value for this attribute is true.
    pub save_external_link_values: Option<bool>,

    /// showBorderUnselectedTables (Show Border Unselected Table)
    ///
    /// Specifies a boolean value that indicates whether a border is drawn around unselected tables in the workbook.
    /// A value of 1 or true indicates borders are drawn around unselected tables.
    /// A value of 0 or false indicates borders are not drawn around unselected tables.
    /// The default value for this attribute is true.
    pub show_border_unselected_tables: Option<bool>,

    /// showInkAnnotation (Show Ink Annotations)
    ///
    /// Specifies a boolean value that indicates whether the book shows ink annotations.
    /// A value of 1 or true indicates that ink annotations are shown in the workbook.
    /// A value of 0 or false indicates that ink annotations are not shown in the workbook.
    /// The default value for this attribute is true.
    pub show_ink_annotation: Option<bool>,

    /// showObjects (Show Objects)
    ///
    /// Specifies how the application shows embedded objects in the workbook.
    /// Possible values:  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.objectdisplayvalues?view=openxml-3.0.1.
    /// The default value for this attribute is "all".
    pub show_objects: Option<String>,

    /// showPivotChartFilter (Show Pivot Chart Filter)
    ///
    /// Specifies a boolean value that indicates whether filtering options are shown for pivot charts in the workbook.
    /// A value of 1 or true indicates filtering options shall be shown for pivot charts.
    /// A value of 0 or false indicates filtering options shall not be shown.
    /// The default value for this attribute is false.
    pub show_pivot_chart_filter: Option<bool>,

    /// updateLinks (Update Links Behavior)
    ///
    /// Specifies how the application updates external links when the workbook is opened.
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.updatelinksbehaviorvalues?view=openxml-3.0.1
    /// The default value for this attribute is userSet.
    pub update_links: Option<String>,
}

impl XlsxWorkbookProperties {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut properties = Self {
            allow_refresh_query: Some(false),
            auto_compress_pictures: Some(true),
            backup_file: Some(false),
            check_compatibility: Some(false),
            code_name: None,
            date1904: Some(false),
            date_compatibility: Some(true),
            default_theme_version: None,
            filter_privacy: Some(false),
            hide_pivot_field_list: Some(false),
            prompted_solutions: Some(false),
            publish_items: Some(false),
            refresh_all_connections: Some(false),
            save_external_link_values: Some(true),
            show_border_unselected_tables: Some(true),
            show_ink_annotation: Some(true),
            show_objects: Some("all".to_owned()),
            show_pivot_chart_filter: Some(false),
            update_links: Some("userSet".to_owned()),
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"allowRefreshQuery" => {
                            properties.allow_refresh_query = string_to_bool(&string_value);
                        }
                        b"autoCompressPictures" => {
                            properties.auto_compress_pictures = string_to_bool(&string_value);
                        }
                        b"backupFile" => {
                            properties.backup_file = string_to_bool(&string_value);
                        }
                        b"checkCompatibility" => {
                            properties.check_compatibility = string_to_bool(&string_value);
                        }
                        b"codeName" => {
                            properties.code_name = Some(string_value);
                        }
                        b"date1904" => {
                            properties.date1904 = string_to_bool(&string_value);
                        }
                        b"dateCompatibility" => {
                            properties.date_compatibility = string_to_bool(&string_value);
                        }
                        b"defaultThemeVersion" => {
                            properties.default_theme_version = string_to_int(&string_value);
                        }
                        b"filterPrivacy" => {
                            properties.filter_privacy = string_to_bool(&string_value);
                        }
                        b"hidePivotFieldList" => {
                            properties.hide_pivot_field_list = string_to_bool(&string_value);
                        }
                        b"promptedSolutions" => {
                            properties.prompted_solutions = string_to_bool(&string_value);
                        }
                        b"publishItems" => {
                            properties.publish_items = string_to_bool(&string_value);
                        }
                        b"refreshAllConnections" => {
                            properties.refresh_all_connections = string_to_bool(&string_value);
                        }
                        b"saveExternalLinkValues" => {
                            properties.save_external_link_values = string_to_bool(&string_value);
                        }
                        b"showBorderUnselectedTables" => {
                            properties.show_border_unselected_tables =
                                string_to_bool(&string_value);
                        }
                        b"showInkAnnotation" => {
                            properties.show_ink_annotation = string_to_bool(&string_value);
                        }
                        b"showObjects" => {
                            properties.show_objects = Some(string_value);
                        }
                        b"showPivotChartFilter" => {
                            properties.show_pivot_chart_filter = string_to_bool(&string_value);
                        }
                        b"updateLinks" => {
                            properties.update_links = Some(string_value);
                        }

                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(properties)
    }
}
