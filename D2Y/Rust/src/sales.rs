// 销售订单读取 + 客户经纬度匹配 + 明细文件匹配
// v3.1.3 从 engine.rs 拆出（原文件过大，且这三步本身是一组"输入侧"逻辑）
use std::collections::HashMap;
use calamine::{open_workbook_auto, Reader};
use crate::{candidate_report, cell_str, detect_columns, detect_header_row,
            Diag, Row, COORD_FIELDS, SALES_FIELDS};

pub(crate) fn read_sales(path: &str, diags: &mut Vec<Diag>) -> Result<Vec<Row>, String> {
    let mut wb = open_workbook_auto(path).map_err(|e| format!("{}", e))?;
    let range = wb.worksheet_range_at(0).ok_or("no sheet")?.map_err(|e| format!("{}", e))?;
    let (rows, cols) = range.get_size();

    // 表头行探测：前 15 行里命中字段最多的一行（至少 3 个字段），失败即报错
    let hr = match detect_header_row(&range, rows, cols, SALES_FIELDS, 3, 15) {
        Some(r) => r,
        None => {
            let first = (0..rows.min(3))
                .map(|r| format!("第{}行: {}", r + 1,
                    (0..cols.min(12)).map(|c| cell_str(&range, r, c)).filter(|s| !s.is_empty())
                        .collect::<Vec<_>>().join(" | ")))
                .collect::<Vec<_>>().join("；");
            diags.push(Diag::error("表头", "销售订单表头行识别失败",
                format!("前 15 行内没有找到同时包含客户单位/订单号等字段的表头行。前几行内容: {}", first)));
            return Err("销售订单表头识别失败".into());
        }
    };
    if hr > 0 {
        diags.push(Diag::warn("表头", &format!("表头位于第 {} 行（通常在第 1 行）", hr + 1),
            "请确认该行确实是表头".to_string()));
    }

    let map = detect_columns(&range, hr, cols, rows, SALES_FIELDS, "销售订单", diags);
    let g = |k: &str, r: usize| map.get(k).map(|&c| cell_str(&range, r, c)).unwrap_or_default();

    let mut result = Vec::new();
    for r in (hr + 1)..rows {
        let cu = g("cu", r); let ad = g("ad", r); let on = g("on", r);
        if cu.is_empty() && ad.is_empty() && on.is_empty() { continue; }
        let ct = g("ct", r);
        let loc = if let Some(p) = cu.find('★') { cu[p+3..].trim().to_string() }
                  else if let Some(p) = cu.find(' ') { cu[p+1..].trim().to_string() }
                  else { cu.clone() };
        let (ph, cn) = if let Some(p) = ct.rfind('-') {
            let ph = ct[p+1..].trim();
            (if ph.len()==11 && ph.starts_with('1') {ph.to_string()} else {String::new()}, ct[..p].trim().to_string())
        } else { (String::new(), ct) };
        let rm = g("rm", r);
        let rm2 = if rm.trim().is_empty() { String::new() } else {
            rm.lines().filter(|l| !l.trim().is_empty()).map(|l| format!("备注：{}", l.trim())).collect::<Vec<_>>().join("\n")
        };
        result.push(Row {
            customer_name: cu, address: ad, order_no: on, phone: ph,
            contact_person: cn, remark: rm2, location_name: loc,
            coords: String::new(), settle_status: g("ss", r), unpaid_amount: 0.0,
        });
    }
    Ok(result)
}

pub(crate) fn match_coords(path: &str, rows: &mut [Row], diags: &mut Vec<Diag>) -> Vec<String> {
    let mut wb = match open_workbook_auto(path) {
        Ok(w) => w,
        Err(_) => {
            diags.push(Diag::warn("客户经纬度", "客户经纬度文件无法读取",
                format!("{} 不存在或无法打开，所有客户的精准坐标将为空", path)));
            return vec![];
        }
    };
    let range = match wb.worksheet_range_at(0) { Some(Ok(r)) => r, _ => {
        diags.push(Diag::warn("客户经纬度", "客户经纬度文件没有工作表", path.to_string()));
        return vec![];
    } };
    let (rn, cols) = range.get_size();
    let map = detect_columns(&range, 0, cols, rn, COORD_FIELDS, "客户经纬度", diags);
    let (a, b) = match (map.get("cu").copied(), map.get("coord").copied()) {
        (Some(a), Some(b)) => (a, b),
        _ => {
            diags.push(Diag::warn("客户经纬度", "缺少客户单位/地图经纬度列，跳过经纬度匹配",
                candidate_report(&range, 0, cols, rn, &COORD_FIELDS[1])));
            return vec![];
        }
    };
    // 提取客户单位中的代码（第一个空格/★前的字母数字）
    fn extract_code(name: &str) -> String {
        let prefix = name.split(|c: char| c == ' ' || c == '★').next().unwrap_or("");
        prefix.chars().filter(|c| c.is_alphanumeric()).collect::<String>().to_uppercase()
    }
    // 建立 代码→经纬度 映射（同一代码取第一条）
    let mut code_map = HashMap::new();
    for r in 1..rn {
        let n = cell_str(&range, r, a); let c = cell_str(&range, r, b);
        if !n.is_empty() && !c.is_empty() {
            let code = extract_code(&n);
            if !code.is_empty() { code_map.entry(code).or_insert(c.clone()); }
        }
    }
    let mut un = Vec::new();
    for row in rows.iter_mut() {
        let code = extract_code(&row.customer_name);
        if let Some(c) = code_map.get(&code) {
            row.coords = c.clone();
        } else if !row.customer_name.trim().is_empty() {
            un.push(row.customer_name.clone());
        }
    }
    un
}

pub(crate) fn match_details(dir: &str, rows: &mut [Row], diags: &mut Vec<Diag>) -> (usize, usize) {
    let mut fm = HashMap::new(); let mut t = 0;
    if let Ok(e) = std::fs::read_dir(dir) {
        for en in e.flatten() {
            let n = en.file_name().to_string_lossy().to_string();
            if (n.ends_with(".xls") || n.ends_with(".xlsx")) && !n.starts_with("~$") {
                let nums: String = n.chars().filter(|c| c.is_ascii_digit()).collect();
                if !nums.is_empty() { fm.insert(nums, en.path().to_string_lossy().to_string()); t += 1; }
            }
        }
    }
    if fm.is_empty() {
        diags.push(Diag::warn("明细匹配", "明细输出目录内没有可用文件",
            format!("{} 下没有 .xls/.xlsx，无法为任何订单附加明细", dir)));
    }
    let mut m = 0;
    for row in rows.iter_mut() {
        if row.order_no.is_empty() { continue; }
        let f = fm.iter().find(|(n, _)| row.order_no.contains(*n) || n.contains(&row.order_no));
        if let Some((_, p)) = f {
            if let Ok(mut wb) = open_workbook_auto(p) {
                if let Some(Ok(r)) = wb.worksheet_range_at(0) {
                    if r.get_size().0 > 2 && r.get_size().1 > 19 {
                        let t3 = cell_str(&r, 2, 19);
                        if !t3.is_empty() { row.remark = format!("{}\n\n{}", t3, row.remark); m += 1; }
                    }
                }
            }
        }
    }
    (t, m)
}
