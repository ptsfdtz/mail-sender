use calamine::{open_workbook_auto, Data, Reader};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Clone, Copy, Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum BatchType {
    Interview,
    Admission,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Recipient {
    pub row: usize,
    pub name: String,
    pub email: String,
    pub department: Option<String>,
    pub qq_group_id: Option<String>,
    pub outcome: Outcome,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum Outcome {
    Interview,
    Admitted,
    Rejected,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImportResult {
    pub recipients: Vec<Recipient>,
    pub skipped_rows: Vec<usize>,
}

fn text(cell: Option<&Data>) -> String {
    match cell {
        Some(Data::String(value)) => value.trim().to_owned(),
        Some(Data::Float(value)) if value.fract() == 0.0 => format!("{value:.0}"),
        Some(Data::Float(value)) => value.to_string(),
        Some(Data::Int(value)) => value.to_string(),
        Some(Data::Bool(value)) => value.to_string(),
        Some(value) => value.to_string().trim().to_owned(),
        None => String::new(),
    }
}

pub fn read_workbook(path: &str, batch_type: BatchType) -> Result<ImportResult, String> {
    let mut workbook =
        open_workbook_auto(path).map_err(|error| format!("无法打开 Excel：{error}"))?;
    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| "Excel 中没有工作表".to_string())?
        .map_err(|error| format!("无法读取工作表：{error}"))?;
    let mut rows = range.rows();
    let headers = rows.next().ok_or_else(|| "Excel 是空表".to_string())?;
    let columns: HashMap<String, usize> = headers
        .iter()
        .enumerate()
        .map(|(index, cell)| (text(Some(cell)), index))
        .collect();

    let required = match batch_type {
        BatchType::Interview => vec!["姓名", "邮箱"],
        BatchType::Admission => vec!["姓名", "邮箱", "录取部门", "群号"],
    };
    for header in required {
        if !columns.contains_key(header) {
            return Err(format!("Excel 缺少“{header}”列"));
        }
    }

    let mut recipients = Vec::new();
    let mut skipped_rows = Vec::new();
    for (offset, row) in rows.enumerate() {
        let row_number = offset + 2;
        let get = |name: &str| text(columns.get(name).and_then(|index| row.get(*index)));
        let name = get("姓名");
        let email = get("邮箱");
        if email.is_empty() {
            if !name.is_empty() || row.iter().any(|cell| !text(Some(cell)).is_empty()) {
                skipped_rows.push(row_number);
            }
            continue;
        }
        let department = get("录取部门");
        let qq_group_id = get("群号");
        let outcome = match batch_type {
            BatchType::Interview => Outcome::Interview,
            BatchType::Admission if !department.is_empty() && department != "无" => {
                Outcome::Admitted
            }
            BatchType::Admission => Outcome::Rejected,
        };
        recipients.push(Recipient {
            row: row_number,
            name,
            email,
            department: (!department.is_empty()).then_some(department),
            qq_group_id: (!qq_group_id.is_empty()).then_some(qq_group_id),
            outcome,
        });
    }
    if recipients.is_empty() {
        return Err("名单中没有包含邮箱的有效收件人".into());
    }
    Ok(ImportResult {
        recipients,
        skipped_rows,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    #[test]
    fn numeric_cells_keep_readable_values() {
        assert_eq!(text(Some(&Data::Float(12345.0))), "12345");
        assert_eq!(text(Some(&Data::Int(678))), "678");
    }

    #[test]
    fn legacy_empty_workbooks_report_a_clear_error() {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("..")
            .join("data");
        let interview = read_workbook(
            root.join("面试名单.xlsx").to_str().unwrap(),
            BatchType::Interview,
        )
        .unwrap_err();
        let admission = read_workbook(
            root.join("录取名单.xlsx").to_str().unwrap(),
            BatchType::Admission,
        )
        .unwrap_err();
        assert_eq!(interview, "名单中没有包含邮箱的有效收件人");
        assert_eq!(admission, "名单中没有包含邮箱的有效收件人");
    }
}
