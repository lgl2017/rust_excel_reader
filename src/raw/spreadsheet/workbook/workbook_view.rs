use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_int, string_to_unsignedint},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.bookviews?view=openxml-3.0.1
///
/// This element specifies the collection of workbook views of the enclosing workbook.
/// Each view can specify a window position, filter options, and other configurations.
/// There is no limit on the number of workbook views that can be defined for a workbook.
///
/// Example:
/// ```
/// <bookViews>
///     <workbookView xWindow="120" yWindow="45" windowWidth="15135" windowHeight="7650" activeTab="4"/>
/// </bookViews>
/// ```
/// bookViews (Workbook Views)
pub type WorkbookViews = Vec<WorkbookView>;

pub(crate) fn load_bookviews(reader: &mut XmlReader) -> anyhow::Result<WorkbookViews> {
    let mut views: WorkbookViews = vec![];

    let mut buf = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"workbookView" => {
                views.push(WorkbookView::load(reader, e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"bookViews" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(views)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookview?view=openxml-3.0.1
///
/// This element specifies a single Workbook view.
/// Units for window widths and other dimensions are expressed in twips.
/// Twip measurements are portable between different display resolutions.
/// The formula is (screen pixels) * (20 * 72) / (logical device dpi), where the logical device dpi can be different for x and y coordinates.
///
/// Example
/// ```
/// <workbookView xWindow="120" yWindow="45" windowWidth="15135" windowHeight="7650" activeTab="4"/>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct WorkbookView {
    // extLst (Future Feature Data Storage Area) Not supported

    // Attributes
    /// activeTab (Active Sheet Index)
    ///
    /// Specifies an unsignedInt that contains the index to the active sheet in this book view.
    /// The default value for this attribute is 0.
    pub active_tab: Option<u64>,

    /// autoFilterDateGrouping (AutoFilter Date Grouping)
    ///
    /// Specifies a boolean value that indicates whether to group dates when presenting the user with filtering options in the user interface.
    /// A value of 1 or true indicates that dates are grouped.
    /// A value of 0 or false indicates that dates are not grouped.
    /// The default value for this attribute is true.
    pub auto_filter_date_grouping: Option<bool>,

    /// firstSheet (First Sheet)
    ///
    /// Specifies the index to the first sheet in this book view.
    /// The default value for this attribute is 0.
    pub first_sheet: Option<u64>,

    /// minimized (Minimized)
    ///
    /// Specifies a boolean value that indicates whether the workbook window is minimized.
    /// A value of 1 or true indicates the workbook window is minimized.
    /// A value of 0 or false indicates the workbook window is not minimized.
    /// The default value for this attribute is false.
    pub minimized: Option<bool>,

    /// showHorizontalScroll (Show Horizontal Scroll)
    ///
    /// Specifies a boolean value that indicates whether to display the horizontal scroll bar in the user interface.
    /// A value of 1 or true indicates that the horizontal scrollbar shall be shown.
    /// A value of 0 or false indicates that the horizontal scrollbar shall not be shown.
    /// The default value for this attribute is true.
    pub show_horizontal_scroll: Option<bool>,

    /// showSheetTabs (Show Sheet Tabs)
    ///
    ///Specifies a boolean value that indicates whether to display the sheet tabs in the user interface.
    /// A value of 1 or true indicates that sheet tabs shall be shown.
    /// A value of 0 or false indicates that sheet tabs shall not be shown.
    /// The default value for this attribute is true.
    pub show_sheet_tabs: Option<bool>,

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
    pub tab_ratio: Option<u64>,

    /// visibility (Visibility)	Specifies visible state of the workbook window.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.visibilityvalues?view=openxml-3.0.1.
    /// The default value for this attribute is "visible."
    pub visibility: Option<String>,

    /// windowHeight (Window Height)
    ///
    /// Specifies the height of the workbook window.
    /// The unit of measurement for this value is twips.
    pub window_height: Option<u64>,

    /// windowWidth (Window Width)
    ///
    /// Specifies the width of the workbook window.
    /// The unit of measurement for this value is twips.
    pub window_width: Option<u64>,

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

impl WorkbookView {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut view = Self {
            active_tab: Some(0),
            auto_filter_date_grouping: Some(true),
            first_sheet: Some(0),
            minimized: Some(false),
            show_horizontal_scroll: Some(true),
            show_sheet_tabs: Some(true),
            show_vertical_scroll: Some(true),
            tab_ratio: Some(600),
            visibility: Some("visible".to_owned()),
            window_height: None,
            window_width: None,
            x_window: None,
            y_window: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"activeTab" => {
                            view.active_tab = string_to_unsignedint(&string_value);
                        }
                        b"autoFilterDateGrouping" => {
                            view.auto_filter_date_grouping = string_to_bool(&string_value);
                        }
                        b"firstSheet" => {
                            view.first_sheet = string_to_unsignedint(&string_value);
                        }
                        b"minimized" => {
                            view.minimized = string_to_bool(&string_value);
                        }
                        b"showHorizontalScroll" => {
                            view.show_horizontal_scroll = string_to_bool(&string_value);
                        }
                        b"showSheetTabs" => {
                            view.show_sheet_tabs = string_to_bool(&string_value);
                        }
                        b"showVerticalScroll" => {
                            view.show_vertical_scroll = string_to_bool(&string_value);
                        }
                        b"tabRatio" => {
                            view.tab_ratio = string_to_unsignedint(&string_value);
                        }
                        b"visibility" => {
                            view.visibility = Some(string_value);
                        }
                        b"windowHeight" => {
                            view.window_height = string_to_unsignedint(&string_value);
                        }
                        b"windowWidth" => {
                            view.window_width = string_to_unsignedint(&string_value);
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
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"workbookView" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(view)
    }
}
