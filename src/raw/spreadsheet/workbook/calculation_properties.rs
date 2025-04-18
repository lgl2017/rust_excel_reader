use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::{string_to_bool, string_to_float, string_to_int};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculationproperties?view=openxml-3.0.1
///
/// This element defines the collection of properties the application uses to record calculation status and details.
///
/// Example
/// ```
/// <calcPr calcId="122211" calcMode="auto" refMode="R1C1" iterate="1" fullPrecision="0"/>
/// ```
/// calcPr (Calculation Properties)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCalculationProperties {
    /// Calc Completed
    // tag: calcCompleted
    pub calculation_completed: Option<bool>,

    /// Calculation Id
    // tag: calcId
    pub calculation_id: Option<i64>,

    /// Calculation Mode
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calculatemodevalues?view=openxml-3.0.1
    // tag: calcMode
    pub calculation_mode: Option<String>,

    /// Calculate On Save
    // tag: calcOnSave
    pub calculation_on_save: Option<bool>,

    /// Concurrent Calculations
    // tag: concurrentCalc
    pub concurrent_calculation: Option<bool>,

    /// Concurrent Thread Manual Count
    // tag: concurrentManualCount
    pub concurrent_manual_count: Option<i64>,

    /// Force Full Calculation
    // tag: forceFullCalc
    pub force_full_calculation: Option<bool>,

    /// Full Calculation On Load
    // tag: fullCalcOnLoad
    pub full_calculation_on_load: Option<bool>,

    /// Full Precision Calculation
    // tag: fullPrecision
    pub full_precision: Option<bool>,

    /// tag: iterate
    pub iterate: Option<bool>,

    /// Iteration Count
    // tag: iterateCount
    pub iterate_count: Option<i64>,

    /// Iterative Calculation Delta
    // tag: iterateDelta
    pub iterate_delta: Option<f64>,

    /// Reference Mode
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.referencemodevalues?view=openxml-3.0.1
    // tag: refMode
    pub reference_mode: Option<String>,
}

impl XlsxCalculationProperties {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut properties = Self {
            calculation_completed: None,
            calculation_id: None,
            calculation_mode: None,
            calculation_on_save: None,
            concurrent_calculation: None,
            concurrent_manual_count: None,
            force_full_calculation: None,
            full_calculation_on_load: None,
            full_precision: None,
            iterate: None,
            iterate_count: None,
            iterate_delta: None,
            reference_mode: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"calcCompleted" => {
                            properties.calculation_completed = string_to_bool(&string_value);
                        }
                        b"calcId" => {
                            properties.calculation_id = string_to_int(&string_value);
                        }
                        b"calcMode" => {
                            properties.calculation_mode = Some(string_value);
                        }
                        b"calcOnSave" => {
                            properties.calculation_on_save = string_to_bool(&string_value);
                        }
                        b"concurrentCalc" => {
                            properties.concurrent_calculation = string_to_bool(&string_value);
                        }
                        b"concurrentManualCount" => {
                            properties.concurrent_manual_count = string_to_int(&string_value);
                        }
                        b"forceFullCalc" => {
                            properties.force_full_calculation = string_to_bool(&string_value);
                        }
                        b"fullCalcOnLoad" => {
                            properties.full_calculation_on_load = string_to_bool(&string_value);
                        }
                        b"fullPrecision" => {
                            properties.full_precision = string_to_bool(&string_value);
                        }
                        b"iterate" => {
                            properties.iterate = string_to_bool(&string_value);
                        }
                        b"iterateCount" => {
                            properties.iterate_count = string_to_int(&string_value);
                        }
                        b"iterateDelta" => {
                            properties.iterate_delta = string_to_float(&string_value);
                        }
                        b"refMode" => {
                            properties.reference_mode = Some(string_value);
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
