use std::fmt;

#[derive(Debug, Clone)]
pub struct ToolError(pub String);

impl ToolError {
    pub fn new(msg: impl Into<String>) -> Self {
        ToolError(msg.into())
    }
}

impl fmt::Display for ToolError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl std::error::Error for ToolError {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn displays_its_message() {
        let err = ToolError::new("Outlook exploded");
        assert_eq!(err.to_string(), "Outlook exploded");
    }

    #[test]
    fn new_accepts_string_and_str() {
        let a = ToolError::new("literal");
        let b = ToolError::new(String::from("owned"));
        assert_eq!(a.to_string(), "literal");
        assert_eq!(b.to_string(), "owned");
    }
}

impl From<ToolError> for rmcp::ErrorData {
    fn from(err: ToolError) -> Self {
        rmcp::ErrorData::internal_error(err.0, None)
    }
}
