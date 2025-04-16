#[derive(Debug, Clone, PartialEq)]
pub struct Formula {
    pub formula: String,
    pub last_calculated_value: Option<String>,
}
