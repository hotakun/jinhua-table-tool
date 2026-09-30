// Rust 处理引擎 — DLL 版，供 Python 通过 ctypes 调用
// v3.1.3:
//   1) 列识别加固 —— 字段别名表 + 表头多候选探测 + 取消「按列号兜底」+ 候选列非空率诊断
//   2) 诊断输出 —— 返回 JSON 增加 errors/warnings（含列号、列名、非空率、候选列），便于定位错误点
//   3) force 参数 —— 致命问题默认中止（不生成输出文件），force=1 时记 warning 并继续生成
//   4) 拆分：sales.rs（读订单/经纬度/明细匹配）、output.rs（备注重组/模板写出/统计）
#![allow(unused)]

mod output;
mod sales;

use std::ffi::{CStr, CString};
use std::os::raw::c_char;
use std::collections::HashMap;
use calamine::{open_workbook_auto, Data, Reader};
use chrono::Timelike;

use output::{add_customer_info, compute_stats, count_products, merge, process_remarks,
             read_headers, save_output};
use sales::{match_coords, match_details, read_sales};

type ProgressCb = unsafe extern "C" fn(u32, *const c_char);

static mut CALLBACK: Option<ProgressCb> = None;

fn progress(pct: u32, label: &str) {
    unsafe {
        if let Some(cb) = CALLBACK {
            let s = CString::new(label).unwrap_or_default();
            cb(pct, s.as_ptr());
        }
    }
}

pub(crate) const INPUT_DETAIL: &str = r"D:\明细表格";
pub(crate) const OUTPUT_DETAIL: &str = r"D:\明细表格\MXBG2";

// ========== 诊断 ==========

#[derive(Debug, Clone)]
pub(crate) struct Diag {
    level: &'static str,   // "error" | "warning"
    pub(crate) field: String,
    pub(crate) problem: String,
    pub(crate) detail: String,
}

impl Diag {
    pub(crate) fn error(field: &str, problem: &str, detail: String) -> Diag {
        Diag { level: "error", field: field.to_string(), problem: problem.to_string(), detail }
    }
    pub(crate) fn warn(field: &str, problem: &str, detail: String) -> Diag {
        Diag { level: "warning", field: field.to_string(), problem: problem.to_string(), detail }
    }
}

pub(crate) fn has_error(diags: &[Diag]) -> bool {
    diags.iter().any(|d| d.level == "error")
}

/// JSON 字符串转义（含换行/制表/控制字符，避免诊断文本破坏 JSON）
pub(crate) fn json_escape(s: &str) -> String {
    let mut o = String::with_capacity(s.len() + 8);
    for ch in s.chars() {
        match ch {
            '"' => o.push_str("\\\""),
            '\\' => o.push_str("\\\\"),
            '\n' => o.push_str("\\n"),
            '\r' => o.push_str("\\r"),
            '\t' => o.push_str("\\t"),
            c if (c as u32) < 0x20 => o.push_str(&format!("\\u{:04x}", c as u32)),
            c => o.push(c),
        }
    }
    o
}

fn diags_json(diags: &[Diag], level: &str) -> String {
    diags.iter()
        .filter(|d| d.level == level)
        .map(|d| format!(
            r#"{{"field":"{}","problem":"{}","detail":"{}"}}"#,
            json_escape(&d.field), json_escape(&d.problem), json_escape(&d.detail)))
        .collect::<Vec<_>>()
        .join(",")
}

// ========== 列识别规格（精确名 → 别名 → 排除词）==========

pub(crate) struct FieldSpec {
    pub(crate) key: &'static str,
    pub(crate) label: &'static str,
    exact: &'static [&'static str],
    alias: &'static [&'static str],
    exclude: &'static [&'static str],
    required: bool,
}

/// 销售订单表
pub(crate) const SALES_FIELDS: &[FieldSpec] = &[
    FieldSpec { key: "cu", label: "客户单位", required: true,
        exact: &["客户单位"], alias: &["客户名称", "单位名称", "门店名称", "客户全称", "客户"],
        exclude: &["编码", "编号", "地址", "电话", "id", "ID"] },
    FieldSpec { key: "on", label: "订单号", required: true,
        exact: &["订单号", "销售单号"], alias: &["订单编号", "销售订单号", "单据编号", "单号"],
        exclude: &["外部", "物流", "快递", "第三方", "平台", "线上", "售后", "退货"] },
    FieldSpec { key: "ad", label: "地址", required: false,
        exact: &["地址"], alias: &["详细地址", "收货地址", "客户地址"], exclude: &["备注"] },
    FieldSpec { key: "ct", label: "联系信息", required: false,
        exact: &["联系信息"], alias: &["联系方式", "联系人信息", "联系电话"], exclude: &[] },
    FieldSpec { key: "rm", label: "备注", required: false,
        exact: &["备注"], alias: &["订单备注", "客户备注", "备注信息"], exclude: &["地址"] },
    FieldSpec { key: "ss", label: "结账状态", required: false,
        exact: &["结账状态", "结算状态"], alias: &["付款状态", "支付状态", "结算"], exclude: &[] },
];

/// 明细表（表头在第 2 行，index 1）
pub(crate) const DETAIL_FIELDS: &[FieldSpec] = &[
    FieldSpec { key: "name", label: "商品名称", required: true,
        exact: &["商品名称"], alias: &["商品"], exclude: &["编码", "编号"] },
    FieldSpec { key: "spec", label: "规格", required: false,
        exact: &["规格"], alias: &[], exclude: &[] },
    FieldSpec { key: "unit", label: "单位", required: false,
        exact: &["单位"], alias: &[], exclude: &[] },
    FieldSpec { key: "price", label: "实际价格", required: false,
        exact: &["实际价格"], alias: &["单价", "销售价格"], exclude: &["零售", "批发", "总"] },
    FieldSpec { key: "qty", label: "订货数量", required: true,
        exact: &["订货数量"], alias: &["数量"], exclude: &["出库", "库存", "赠"] },
    FieldSpec { key: "amount", label: "金额", required: false,
        exact: &["金额"], alias: &[], exclude: &["缺货", "分摊", "总", "税"] },
    FieldSpec { key: "remark", label: "备注", required: false,
        exact: &["备注"], alias: &[], exclude: &[] },
];

/// 客户经纬度表
pub(crate) const COORD_FIELDS: &[FieldSpec] = &[
    FieldSpec { key: "cu", label: "客户单位", required: true,
        exact: &["客户单位"], alias: &["客户名称", "门店名称", "客户"],
        exclude: &["编码", "编号", "地址", "电话"] },
    FieldSpec { key: "coord", label: "地图经纬度", required: true,
        exact: &["地图经纬度"], alias: &["经纬度", "坐标"], exclude: &[] },
];

/// 某列是否命中字段：先查排除词，再按 pass 决定用精确名还是别名
pub(crate) fn spec_hit(sp: &FieldSpec, s: &str, exact_pass: bool) -> bool {
    if s.is_empty() { return false; }
    if sp.exclude.iter().any(|x| s.contains(*x)) { return false; }
    if exact_pass { sp.exact.iter().any(|n| s == *n) }
    else { sp.alias.iter().any(|n| s.contains(*n)) }
}

/// 该列在数据区的非空率：(非空数, 总行数)
fn col_fill(range: &calamine::Range<Data>, rows: usize, hr: usize, c: usize) -> (usize, usize) {
    let mut ne = 0usize;
    let mut tot = 0usize;
    for r in (hr + 1)..rows {
        tot += 1;
        if !cell_str(range, r, c).is_empty() { ne += 1; }
    }
    (ne, tot)
}

/// 列出该字段的所有同名列及其非空率；被「排除词」挡掉的列标注（已排除），
/// 便于一眼看出「选错列」——例如订单号候选里出现 c2『外部订单号』非空 0/386（已排除）
pub(crate) fn candidate_report(range: &calamine::Range<Data>, hr: usize, cols: usize,
                               rows: usize, sp: &FieldSpec) -> String {
    let mut items = Vec::new();
    for c in 0..cols {
        let s = cell_str(range, hr, c);
        let by_name = sp.exact.iter().any(|n| s == *n) || sp.alias.iter().any(|n| s.contains(*n));
        let excluded = sp.exclude.iter().any(|n| s.contains(*n));
        if by_name || excluded {
            let (ne, tot) = col_fill(range, rows, hr, c);
            let tag = if excluded { "（已排除）" } else { "" };
            items.push(format!("c{}『{}』非空 {}/{}{}", c, s, ne, tot, tag));
        }
    }
    if items.is_empty() { "无同名列".to_string() }
    else { format!("候选列: {}", items.join(" / ")) }
}

/// 表头行探测：先按「精确列名」严格打分（数据行几乎不可能命中列名），找不到再退回
/// 「精确名 + 别名」的模糊打分。取命中字段最多的一行，至少 min_hit 个字段。
pub(crate) fn detect_header_row(range: &calamine::Range<Data>, rows: usize, cols: usize,
                                specs: &[FieldSpec], min_hit: usize, max_scan: usize) -> Option<usize> {
    for strict in [true, false] {
        let mut best: Option<(usize, usize)> = None;
        for r in 0..rows.min(max_scan) {
            let mut score = 0usize;
            for c in 0..cols {
                let s = cell_str(range, r, c);
                let hit = specs.iter().any(|sp| {
                    if strict { sp.exact.iter().any(|n| s == *n) }
                    else { spec_hit(sp, &s, true) || spec_hit(sp, &s, false) }
                });
                if hit { score += 1; }
            }
            if score >= min_hit && best.map_or(true, |(_, bs)| score > bs) {
                best = Some((r, score));
            }
        }
        if best.is_some() { return best.map(|(r, _)| r); }
    }
    None
}

/// 列识别：两趟匹配（精确名 → 别名）+ 已命中字段不覆盖 + 诊断
pub(crate) fn detect_columns(range: &calamine::Range<Data>, hr: usize, cols: usize, rows: usize,
                             specs: &[FieldSpec], sheet: &str,
                             diags: &mut Vec<Diag>) -> HashMap<&'static str, usize> {
    let mut map: HashMap<&'static str, usize> = HashMap::new();
    for pass in 0..2u8 {
        for c in 0..cols {
            let s = cell_str(range, hr, c);
            for sp in specs {
                if map.contains_key(sp.key) { continue; }
                if spec_hit(sp, &s, pass == 0) { map.insert(sp.key, c); }
            }
        }
    }
    for sp in specs {
        match map.get(sp.key).copied() {
            None => {
                let detail = candidate_report(range, hr, cols, rows, sp);
                if sp.required {
                    diags.push(Diag::error(sp.label,
                        &format!("{} 未找到「{}」列", sheet, sp.label), detail));
                } else {
                    diags.push(Diag::warn(sp.label,
                        &format!("{} 未找到「{}」列，该字段将为空", sheet, sp.label), detail));
                }
            }
            Some(c) => {
                // 只对必需字段（客户单位/订单号）做「非空率偏低」检查：
                // 备注、结账状态等字段本来就可能大量为空，对它们检查会产生误报
                let (ne, tot) = col_fill(range, rows, hr, c);
                if sp.required && tot >= 5 && ne * 100 / tot < 50 {
                    diags.push(Diag::warn(sp.label,
                        &format!("{}{} 列疑似选错（非空率偏低）", sheet, sp.label),
                        format!("当前选用 c{}『{}』非空 {}/{}；{}",
                            c, cell_str(range, hr, c), ne, tot,
                            candidate_report(range, hr, cols, rows, sp))));
                }
            }
        }
    }
    map
}

// ========== 导出函数 ==========

#[no_mangle]
pub extern "C" fn engine_process(
    delete: u8,
    stats: u8,
    force: u8,
    cb: ProgressCb,
) -> *mut c_char {
    unsafe { CALLBACK = Some(cb); }
    let json = run_all(delete != 0, stats != 0, force != 0);
    CString::new(json).unwrap_or_default().into_raw()
}

#[no_mangle]
pub extern "C" fn engine_free(ptr: *mut c_char) {
    if !ptr.is_null() {
        unsafe { drop(CString::from_raw(ptr)); }
    }
}

/// 统一的结果 JSON：err 非空表示失败；fatal=true 表示因致命问题中止且未生成文件
fn build_json(err: Option<&str>, fname: &str, orders: usize, customers: usize, matched: usize,
              products: i64, no_coords: &[String], path: &str,
              diags: &[Diag], fatal: bool) -> String {
    let mut s = String::from("{");
    if let Some(e) = err {
        s.push_str(&format!("\"error\":\"{}\",", json_escape(e)));
    }
    let nc: Vec<String> = no_coords.iter().map(|n| format!("\"{}\"", json_escape(n))).collect();
    s.push_str(&format!(
        r#""file":"{}","orders":{},"customers":{},"matched":{},"products":{},"no_coords":[{}],"path":"{}","fatal":{},"errors":[{}],"warnings":[{}]}}"#,
        json_escape(fname), orders, customers, matched, products,
        nc.join(","), json_escape(path), fatal,
        diags_json(diags, "error"), diags_json(diags, "warning")));
    s
}

fn run_all(delete_files: bool, do_stats: bool, force: bool) -> String {
    for d in &[r"D:\订单表格", INPUT_DETAIL, OUTPUT_DETAIL] {
        let _ = std::fs::create_dir_all(d);
    }
    let mut diags: Vec<Diag> = Vec::new();

    progress(10, "正在处理明细表格...");
    if let Err(e) = run_detail(&mut diags) {
        diags.push(Diag::error("明细表格", "明细处理失败", e.clone()));
        return build_json(Some(&e), "", 0, 0, 0, 0, &[], "", &diags, true);
    }
    progress(30, "明细表格处理完成");

    if has_error(&diags) && !force {
        let msg = "明细表格列识别未通过，已中止（详见诊断）";
        return build_json(Some(msg), "", 0, 0, 0, 0, &[], "", &diags, true);
    }

    progress(40, "正在处理订单表格...");
    let (no_coords, fname, output_path, orders, customers, matched, products) =
        match run_order(&mut diags, force) {
            Ok(v) => v,
            Err(e) => {
                let fatal = !force && has_error(&diags);
                return build_json(Some(&e), "", 0, 0, 0, 0, &[], "", &diags, fatal);
            }
        };

    if do_stats {
        progress(72, "正在统计商品明细...");
        let (lines, total) = compute_stats();
        if !lines.is_empty() {
            let escaped: Vec<String> = lines.iter()
                .map(|l| format!("\"{}\"", json_escape(l)))
                .collect();
            progress(75, &format!("STATS:{}|{}", total, escaped.join(",")));
        }
    }

    if delete_files {
        progress(85, "正在清空明细表格目录...");
        if let Ok(entries) = std::fs::read_dir(INPUT_DETAIL) {
            for e in entries.flatten() {
                let p = e.path();
                let _ = if p.is_dir() { std::fs::remove_dir_all(&p) } else { std::fs::remove_file(&p) };
            }
        }
    }

    progress(100, "执行完毕");
    build_json(None, &fname, orders, customers, matched, products,
               &no_coords, &output_path, &diags, false)
}

// ========== 明细处理 ==========

fn short_name(p: &str) -> String {
    std::path::Path::new(p).file_name()
        .map(|n| n.to_string_lossy().to_string()).unwrap_or_else(|| p.to_string())
}

fn run_detail(diags: &mut Vec<Diag>) -> Result<(), String> {
    let dir = std::path::Path::new(INPUT_DETAIL);
    if !dir.exists() { return Err("明细目录不存在".into()); }
    std::fs::create_dir_all(OUTPUT_DETAIL).map_err(|e| e.to_string())?;

    let mut files = Vec::new();
    for e in std::fs::read_dir(dir).map_err(|e| e.to_string())? {
        let e = e.map_err(|e| e.to_string())?;
        let n = e.file_name().to_string_lossy().to_string();
        if (n.ends_with(".xls") || n.ends_with(".xlsx")) && !n.starts_with("~$") {
            files.push(e.path().to_string_lossy().to_string());
        }
    }
    if files.is_empty() {
        diags.push(Diag::warn("明细表格", "明细目录内没有 Excel 文件",
            format!("{} 下未找到 .xls/.xlsx（订单将没有明细可匹配）", INPUT_DETAIL)));
    }
    for f in &files {
        if let Err(e) = process_single_detail(f, diags) {
            diags.push(Diag::warn("明细表格", "明细文件处理失败",
                format!("{}: {}", short_name(f), e)));
        }
    }
    Ok(())
}

fn process_single_detail(input: &str, diags: &mut Vec<Diag>) -> Result<(), String> {
    let mut wb = open_workbook_auto(input).map_err(|e| format!("{}", e))?;
    let range = wb.worksheet_range_at(0).ok_or("no sheet")?.map_err(|e| format!("{}", e))?;
    let (rows, cols) = range.get_size();
    if rows < 3 { return Ok(()); }

    // 明细表头固定在第 2 行（index 1）
    let hr = 1usize;
    let mut ci: HashMap<&'static str, usize> = HashMap::new();
    for pass in 0..2u8 {
        for c in 0..cols {
            let s = cell_str(&range, hr, c);
            for sp in DETAIL_FIELDS {
                if ci.contains_key(sp.key) { continue; }
                if spec_hit(sp, &s, pass == 0) { ci.insert(sp.key, c); }
            }
        }
    }
    // 必需列校验（不再退化为「取 0 号列」）
    let missing: Vec<&str> = DETAIL_FIELDS.iter()
        .filter(|sp| sp.required && !ci.contains_key(sp.key))
        .map(|sp| sp.label).collect();
    if !missing.is_empty() {
        let detail = DETAIL_FIELDS.iter()
            .map(|sp| format!("{}: {}", sp.label, candidate_report(&range, hr, cols, rows, sp)))
            .collect::<Vec<_>>().join("；");
        diags.push(Diag::error("明细表格",
            &format!("明细文件缺少必需列: {}", missing.join("、")),
            format!("{} | {}", short_name(input), detail)));
        return Err(format!("缺少必需列: {}", missing.join("、")));
    }
    // 其余缺列仅提示
    let missing_opt: Vec<&str> = DETAIL_FIELDS.iter()
        .filter(|sp| !sp.required && !ci.contains_key(sp.key))
        .map(|sp| sp.label).collect();
    if !missing_opt.is_empty() {
        diags.push(Diag::warn("明细表格",
            &format!("明细文件缺少列: {}（相关内容将为空）", missing_opt.join("、")),
            format!("{} | {}", short_name(input),
                candidate_report(&range, hr, cols, rows, &DETAIL_FIELDS[0]))));
    }
    // 备注列保底：找不到就用最后一列（保持原有行为）
    if !ci.contains_key("remark") && cols > 0 { ci.insert("remark", cols - 1); }

    let cell = |r: usize, k: &str| -> String {
        match ci.get(k) { Some(&c) => cell_str(&range, r, c), None => String::new() }
    };

    let mut remarks = Vec::new();
    for r in 2..rows {
        let name = cell(r, "name");
        let spec = cell(r, "spec");
        let unit = cell(r, "unit");
        let price = fmt_price(&cell(r, "price"));
        let qty = fmt_qty(&cell(r, "qty"));
        let amount = fmt_amount(&cell(r, "amount"));
        remarks.push(if name.contains("赠品") {
            format!("{} {} {}：{}；", name, spec, unit, qty)
        } else {
            format!("{} {} {}：{}*{}= {}；", name, spec, unit, price, qty, amount)
        });
    }

    let rc = *ci.get("remark").unwrap_or(&(cols.saturating_sub(1)));
    let nums: String = std::path::Path::new(input).file_name().unwrap().to_string_lossy()
        .chars().filter(|c| c.is_ascii_digit()).collect();
    write_xlsx(&format!("{}\\{}.xlsx", OUTPUT_DETAIL, nums), &range, rows, cols, rc, &remarks)
}

fn write_xlsx(path: &str, range: &calamine::Range<Data>, rows: usize, cols: usize,
              rc: usize, remarks: &[String]) -> Result<(), String> {
    let mut wb = rust_xlsxwriter::Workbook::new();
    let sh = wb.add_worksheet();
    for r in 0..rows {
        for c in 0..cols {
            let v = cell_str(range, r, c);
            if !v.is_empty() { let _ = sh.write_string(r as u32, c as u16, &v); }
        }
        if r >= 2 && r - 2 < remarks.len() {
            let _ = sh.write_string(r as u32, rc as u16, &remarks[r - 2]);
        }
    }
    if rows > 2 { let _ = sh.write_string(2, 19, &remarks.join("\n")); }
    wb.save(path).map_err(|e| format!("{}", e))
}

// ========== 订单处理 ==========

fn run_order(diags: &mut Vec<Diag>, force: bool)
    -> Result<(Vec<String>, String, String, usize, usize, usize, i64), String> {
    let sales_path = r"D:\订单表格\销售订单.xls";
    let coords = r"D:\订单表格\客户经纬度.xls";
    let tmpl = r"D:\订单表格\优路达导入模板.xlsx";
    let detail = r"D:\明细表格\MXBG2";
    let out = r"D:\订单表格";

    for (p, n) in &[(sales_path, "销售订单"), (tmpl, "模板")] {
        if !std::path::Path::new(p).exists() {
            diags.push(Diag::error("文件", &format!("缺少 {}", n), format!("{} 不存在", p)));
            return Err(format!("缺少: {}", n));
        }
    }

    progress(42, "步骤1: 读取订单...");
    let mut rows = read_sales(sales_path, diags)?;
    if has_error(diags) && !force {
        return Err("销售订单列识别未通过，已中止（详见诊断）".into());
    }
    if rows.is_empty() {
        diags.push(Diag::error("销售订单", "没有读到任何数据行",
            format!("{} 表头之下没有有效数据", sales_path)));
        return Err("无数据".into());
    }

    progress(45, "步骤2: 经纬度匹配...");
    let unmatched = match_coords(coords, &mut rows, diags);

    progress(50, "步骤3: 明细匹配...");
    let (_dt, matched) = match_details(detail, &mut rows, diags);

    progress(55, "步骤4: 合并客户...");
    let orders = rows.len();

    // 提前中止：0 匹配说明明细完全没接上（最常见的静默故障）
    if orders > 0 && matched == 0 {
        let detail_txt = format!(
            "处理 {} 单，但没有一单匹配到 {} 的明细文件；{}",
            orders, detail,
            "请先核对「订单号」列是否选对（见上方订单号相关诊断），以及明细是否已生成到 MXBG2");
        if force {
            diags.push(Diag::warn("明细匹配", "匹配明细 0 单（已按你的要求继续生成）", detail_txt));
        } else {
            diags.push(Diag::error("明细匹配", "匹配明细 0 单", detail_txt));
            return Err("匹配明细 0 单，已中止（可用「仍然生成」继续）".into());
        }
    } else if matched < orders {
        diags.push(Diag::warn("明细匹配",
            &format!("匹配明细少于订单数（{}/{}）", matched, orders),
            "部分订单没有对应的明细文件".to_string()));
    }

    rows = merge(rows);
    let customers = rows.len();

    let no_coords: Vec<String> = rows.iter()
        .filter(|r| r.coords.is_empty() && !r.customer_name.is_empty())
        .map(|r| r.location_name.clone()).collect();

    if customers > 0 && rows.iter().all(|r| r.order_no.is_empty()) {
        let txt = "合并后所有订单号都为空，输出文件「订单号」列将整列空白".to_string();
        if force { diags.push(Diag::warn("订单号", "输出的订单号列全为空", txt)); }
        else {
            diags.push(Diag::error("订单号", "输出的订单号列全为空", txt));
            return Err("订单号列为空，已中止（可用「仍然生成」继续）".into());
        }
    }

    progress(60, "步骤5-6: 处理备注...");
    process_remarks(&mut rows);
    add_customer_info(&mut rows);

    progress(68, "步骤7: 生成模板...");
    let headers = read_headers(tmpl)?;

    progress(70, "步骤8: 保存...");
    let opath = save_output(&rows, &headers, out, diags, force)?;
    let fname = std::path::Path::new(&opath).file_name()
        .map(|n| n.to_string_lossy().to_string()).unwrap_or_default();

    let products = count_products(detail);

    // 经纬度未匹配提示
    if !unmatched.is_empty() {
        let ids: Vec<String> = unmatched.iter().take(8).cloned().collect();
        diags.push(Diag::warn("精准坐标",
            &format!("{} 个客户未匹配到经纬度", unmatched.len()),
            format!("例如: {}{}", ids.join("、"),
                if unmatched.len() > 8 { " …" } else { "" })));
    }

    Ok((no_coords, fname, opath, orders, customers, matched, products))
}

// ========== 工具函数 ==========

pub(crate) fn cell_str(range: &calamine::Range<Data>, r: usize, c: usize) -> String {
    match range.get_value((r as u32, c as u32)) {
        Some(Data::String(s)) => s.trim().to_string(),
        Some(Data::Float(f)) => { let s = format!("{:.10}", *f); s.trim_end_matches('0').trim_end_matches('.').to_string() }
        Some(Data::Int(i)) => i.to_string(),
        Some(Data::Bool(b)) => b.to_string(),
        _ => String::new(),
    }
}

pub(crate) fn dw(s: &str) -> usize { s.chars().fold(0, |w, c| w + if c as u32 > 0x2E80 { 4 } else { 2 }) }
pub(crate) fn pad(s: &str, w: usize) -> String { format!("{}{}", s, " ".repeat(w.saturating_sub(dw(s)))) }

fn fmt_price(s: &str) -> String {
    let f = match s.parse::<f64>() { Ok(v) => v, Err(_) => return s.to_string() };
    let s = f.to_string();
    if let Some(d) = s.find('.') { if s[d+1..].len() > 2 { return format!("{:.2}", f); } s.trim_end_matches('0').trim_end_matches('.').to_string() } else { s }
}
fn fmt_qty(s: &str) -> String {
    match s.parse::<f64>() {
        Ok(f) if f.fract() == 0.0 => (f as i64).to_string(),
        Ok(f) => f.to_string(), Err(_) => "0".into()
    }
}
fn fmt_amount(s: &str) -> String { match s.parse::<f64>() { Ok(f) => format!("{:.2}", f), Err(_) => "0.00".into() } }

#[derive(Debug, Clone, Default)]
pub(crate) struct Row {
    pub(crate) customer_name: String,
    pub(crate) address: String,
    pub(crate) order_no: String,
    pub(crate) phone: String,
    pub(crate) contact_person: String,
    pub(crate) remark: String,
    pub(crate) location_name: String,
    pub(crate) coords: String,
    pub(crate) settle_status: String,
    pub(crate) unpaid_amount: f64,
}
