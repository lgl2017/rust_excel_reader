use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_int},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customworkbookviews?view=openxml-3.0.1
///
/// This element defines the collection of custom workbook views that are defined for this workbook.
///
/// Example
/// ```
/// <customWorkbookViews>
///   <customWorkbookView name="CustomView" guid="{CE6681F1-E999-414D-8446-68A031534B57}" maximized="1" xWindow="1" yWindow="1" windowWidth="1024" windowHeight="547" activeSheetId="1"/>
/// </customWorkbookViews>
/// ```
pub type XlsxCustomWorkbookViews = Vec<XlsxCustomWorkbookView>;

pub(crate) fn load_custom_bookviews(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxCustomWorkbookViews> {
    let mut views: XlsxCustomWorkbookViews = vec![];

    let mut buf = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"customWorkbookView" => {
                views.push(XlsxCustomWorkbookView::load(reader, e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"customWorkbookViews" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(views)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customworkbookview?view=openxml-3.0.1
///
/// This element specifies a single custom workbook view. A custom workbook view consists of a set of display and print settings that you can name and apply to a workbook.
///
/// Example
/// ```
///     <customWorkbookView name="CustomView" guid="{CE6681F1-E999-414D-8446-68A031534B57}" maximized="1" xWindow="1"    yWindow="1" windowWidth="1024" windowHeight="547" activeSheetId="1"/>
/// ```
/// customWorkbookView (Custom Workbook View)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCustomWorkbookView {
    /// extLst (Future Feature Data Storage Area) Not supported

    /// Attributes

    /// activeSheetId (Active Sheet in Book View)
    ///
    /// Specifies the sheetId of a sheet in the workbook that identifies to a consuming application the default sheet to display.
    /// Corresponds to a sheetId of a sheet in the sheets collection.
    /// This attribute is required.
    pub active_sheet_id: Option<i64>,

    /// autoUpdate (Auto Update)
    ///
    /// Specifies a boolean value that is an instruction that if the workbook is loaded by a spreadsheet application, that spreadsheet application should automatically update changes at the interval specified by the mergeInterval attribute. This is only applicable for shared workbooks (§18.11).
    /// A value of 1 or true is an instruction to the spreadsheet application to update changes at the interval specified in the mergeInterval attribute.
    /// A value of 0 or false is an instruction to the spreadsheet applicationto update changes whenever the spreadsheet application generates SpreadsheetML representing the workbook.
    /// The default value for this attribute is false.
    pub auto_update: Option<bool>,

    /// changesSavedWin (Changes Saved Win)
    /// Specifies a boolean value that instructs a spreadsheet application to overwrite the persisted version of the document with the updated version being persisted. This is only applicable for shared workbooks in automatic update mode.
    /// A value of 1 or true instructs a spreadsheet application to overwrite changes in the persisted version of a shared workbook when conflicts in data are found.
    /// A value of 0 or false instructs a spreadsheet application to not overwrite changes in the persisted version of a shared workbook when conflicts are found.
    /// The default value for this attribute is false.
    pub changes_saved_win: Option<bool>,

    /// guid (Custom View GUID)
    /// Specifies a globally unique identifier (GUID) for this custom view
    pub guid: Option<String>,

    /// includeHiddenRowCol (Include Hidden Rows & Columns)
    /// Specifies a boolean value that indicates whether to include hidden rows, columns, and filter settings in this custom view.
    /// A value of 1 or true indicates that hidden rows, columns, and filter settings are included in this custom view.
    /// A value of 0 or false indicates that hidden rows, columns, and filter settings are not included.
    /// The default value for this attribute is true.
    pub include_hidden_row_col: Option<bool>,

    /// includePrintSettings (Include Print Settings)
    ///
    /// Specifies a boolean value that indicates whether to include print settings in this custom view.
    /// A value of 1 or true indicates that print settings are included in this custom view.
    /// A value of 0 or false indicates print settings are not included in this custom view.
    /// The default value for this attribute is true.
    pub include_print_settings: Option<bool>,

    /// maximized (Maximized)
    ///
    /// Specifies a boolean value that indicates whether the workbook window is maximized.
    /// A value of 1 or true indicates the workbook window is maximized.
    /// A value of 0 or false indicates the workbook window is not maximized.
    /// The default value for this attribute is false.
    pub maximized: Option<bool>,

    /// mergeInterval (Merge Interval)
    ///
    /// Automatic update interval (in minutes).
    /// Only applicable for shared workbooks in automatic update mode.
    pub merge_interval: Option<i64>,

    /// minimized (Minimized)
    ///
    /// Specifies a boolean value that indicates whether the workbook window is minimized.
    /// A value of 1 or true indicates the workbook window is minimized.
    /// A value of 0 or false indicates the workbook window is not minimized.
    /// The default value for this attribute is false.
    pub minimized: Option<bool>,

    /// name (Custom View Name)
    ///
    /// Specifies the name of the custom view.
    /// This attribute is required.
    pub name: Option<String>,

    /// onlySync (Only Synch)
    ///
    /// Specifies a boolean value that indicates, during automatic update, the current user's changes are not saved. The workbook is only updated with other users' changes. Only applicable for shared workbooks in automatic update mode.
    /// A value of 1 or true indicates the current user's changes is not saved during automatic update.
    /// A value of 0 or false indicates the current user's is saved during automatic update.
    /// The default value for this attribute is false.
    pub only_sync: Option<bool>,

    /// personalView (Personal View)
    ///
    /// Specifies a boolean value that indicates that this custom view is a personal view for a shared workbook user. Only applicable for shared workbooks. Personal views allow each user of a shared workbook to store their individual print and filter settings.
    /// A value of 1 or true indicates this custom view is a personal view for a shared workbook user.
    /// A value of 0 or false indicates this view is not a personal view.
    /// The default value for this attribute is false.
    pub personal_view: Option<bool>,

    /// showComments (Show Comments)
    ///
    /// Specifies how comments are displayed in this custom view
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.commentsvalues?view=openxml-3.0.1
    pub show_comments: Option<bool>,

    /// showFormulaBar (Show Formula Bar)
    ///
    /// Specifies a boolean value that indicates whether to display the formula bar in the application user interface.
    /// A value of 1 or true indicates the formula bar is shown in the user interface.
    /// A value of 0 or false indicates the formula bar is not shown in the user interface.
    /// The default value for this attribute is true.
    pub show_formula_bar: Option<bool>,

    /// showHorizontalScroll (Show Horizontal Scroll)
    ///
    /// Specifies a boolean value that indicates whether to display the horizontal scroll bar in the user interface.
    /// A value of 1 or true indicates that the horizontal scrollbar shall be shown.
    /// A value of 0 or false indicates that the horizontal scrollbar shall not be shown.
    /// The default value for this attribute is true.
    pub show_horizontal_scroll: Option<bool>,

    /// showObjects (Show Objects)
    ///
    /// Specifies how the application shows embedded objects in the workbook.
    /// Possible values:  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.objectdisplayvalues?view=openxml-3.0.1.
    /// The default value for this attribute is "all".
    pub show_objects: Option<String>,

    /// showSheetTabs (Show Sheet Tabs)
    ///
    /// Specifies a boolean value that indicates whether to display the sheet tabs in the user interface.
    /// A value of 1 or true indicates that sheet tabs shall be shown.
    /// A value of 0 or false indicates that sheet tabs shall not be shown.
    /// The default value for this attribute is true.
    pub show_sheet_tabs: Option<bool>,

    /// showStatusbar (Show Status Bar)
    ///
    /// Specifies a boolean value that indicates whether to display the status bar in the user interface.
    /// A value of 1 or true indicates that the status bar is shown.
    /// A value of 0 or false indicates the status bar is not shown.
    /// The default value for this attribute is true.
    pub show_statusbar: Option<bool>,

    /// showVerticalScroll (Show Vertical Scroll)
    ///
    /// Specifies a boolean value that indicates whether to display the vertical scroll bar.
    /// A value of 1 or true indicates the vertical scrollbar shall be shown.
    /// A value of 0 or false indicates the vertical scrollbar shall not be shown.
    /// The default value for this attribute is true.
    pub show_vertical_scroll: Option<bool>,

    /// tabRatio (Sheet Tab Ratio)
    ///
    /// Specifies ratio between the workbook tabs bar and the horizontal scroll bar.
    /// The default value for this attribute is 600.
    pub tab_ratio: Option<i64>,

    /// windowHeight (Window Height)
    ///
    /// Specifies the height of the workbook window.
    /// The unit of measurement for this value is twips.
    pub window_height: Option<i64>,

    /// windowWidth (Window Width)
    ///
    /// Specifies the width of the workbook window.
    /// The unit of measurement for this value is twips.
    pub window_width: Option<i64>,

    /// xWindow (Upper Left Corner (X Coordinate))
    ///
    /// Specifies the X coordinate for the upper left corner of the workbook window.
    /// The unit of measurement for this value is twips.
    pub x_window: Option<i64>,

    /// yWindow (Upper Left Corner (Y Coordinate))
    ///
    /// Specifies the Y coordinate for the upper left corner of the workbook window.
    /// The unit of measurement for this value is twips.
    pub y_window: Option<i64>,
}

impl XlsxCustomWorkbookView {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut view = Self {
            minimized: None,
            show_horizontal_scroll: None,
            show_sheet_tabs: None,
            show_vertical_scroll: None,
            tab_ratio: None,
            window_height: None,
            window_width: None,
            x_window: None,
            y_window: None,
            active_sheet_id: None,
            auto_update: None,
            changes_saved_win: None,
            guid: None,
            include_hidden_row_col: None,
            include_print_settings: None,
            maximized: None,
            merge_interval: None,
            name: None,
            only_sync: None,
            personal_view: None,
            show_comments: None,
            show_formula_bar: None,
            show_objects: None,
            show_statusbar: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"activeSheetId" => {
                            view.active_sheet_id = string_to_int(&string_value);
                        }
                        b"autoUpdate" => {
                            view.auto_update = string_to_bool(&string_value);
                        }
                        b"changesSavedWin" => {
                            view.changes_saved_win = string_to_bool(&string_value);
                        }
                        b"guid" => {
                            view.guid = Some(string_value);
                        }
                        b"includeHiddenRowCol" => {
                            view.include_hidden_row_col = string_to_bool(&string_value);
                        }
                        b"includePrintSettings" => {
                            view.include_print_settings = string_to_bool(&string_value);
                        }
                        b"maximized" => {
                            view.maximized = string_to_bool(&string_value);
                        }
                        b"mergeInterval" => {
                            view.merge_interval = string_to_int(&string_value);
                        }
                        b"minimized" => {
                            view.minimized = string_to_bool(&string_value);
                        }
                        b"name" => {
                            view.name = Some(string_value);
                        }
                        b"onlySync" => {
                            view.only_sync = string_to_bool(&string_value);
                        }
                        b"personalView" => {
                            view.personal_view = string_to_bool(&string_value);
                        }
                        b"showComments" => {
                            view.show_comments = string_to_bool(&string_value);
                        }
                        b"showFormulaBar" => {
                            view.show_formula_bar = string_to_bool(&string_value);
                        }
                        b"showHorizontalScroll" => {
                            view.show_horizontal_scroll = string_to_bool(&string_value);
                        }
                        b"showObjects" => {
                            view.show_objects = Some(string_value);
                        }
                        b"showSheetTabs" => {
                            view.show_sheet_tabs = string_to_bool(&string_value);
                        }
                        b"showStatusbar" => {
                            view.show_statusbar = string_to_bool(&string_value);
                        }
                        b"showVerticalScroll" => {
                            view.show_vertical_scroll = string_to_bool(&string_value);
                        }
                        b"tabRatio" => {
                            view.tab_ratio = string_to_int(&string_value);
                        }
                        b"windowHeight" => {
                            view.window_height = string_to_int(&string_value);
                        }
                        b"windowWidth" => {
                            view.window_width = string_to_int(&string_value);
                        }
                        b"xWindow" => {
                            view.x_window = string_to_int(&string_value);
                        }
                        b"yWindow" => {
                            view.y_window = string_to_int(&string_value);
                        }

                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"customWorkbookView" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        if view.active_sheet_id.is_none() || view.name.is_none() {
            bail!("Requried attributes for custom workbook view is unspecified.");
        }

        Ok(view)
    }
}
