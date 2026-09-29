// Rust 处理引擎 — DLL 版，供 Python 通过 ctypes 调用
#![allow(unused)]

use std::ffi::{CStr, CString};
use std::os::raw::c_char;
use std::collections::HashMap;
use calamine::{open_workbook_auto, Data, Reader};
use chrono::Timelike;

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

const INPUT_DETAIL: &str = r"D:\明细表格";
const OUTPUT_DETAIL: &str = r"D:\明细表格\MXBG2";

// ========== 导出函数 ==========

#[no_mangle]
pub extern "C" fn engine_process(
    delete: u8,
    stats: u8,
    cb: ProgressCb,
) -> *mut c_char {
    unsafe { CALLBACK = Some(cb); }

    let result = run_all(delete != 0, stats != 0);
    let json = match result {
        Ok(s) => s,
        Err(e) => format!(r#"{{"error":"{}"}}"#, e.replace('"', "'")),
    };
    CString::new(json).unwrap().into_raw()
}

#[no_mangle]
pub extern "C" fn engine_free(ptr: *mut c_char) {
    if !ptr.is_null() {
        unsafe { drop(CString::from_raw(ptr)); }
    }
}

fn run_all(delete_files: bool, do_stats: bool) -> Result<String, String> {
    for d in &[r"D:\订单表格", INPUT_DETAIL, OUTPUT_DETAIL] {
        let _ = std::fs::create_dir_all(d);
    }

    progress(10, "正在处理明细表格...");
    run_detail()?;
    progress(30, "明细表格处理完成");

    progress(40, "正在处理订单表格...");
    let (no_coords, fname, output_path, orders, customers, matched, products) = run_order()?;

    if do_stats {
        progress(72, "正在统计商品明细...");
        let (lines, total) = compute_stats();
        if !lines.is_empty() {
            let escaped: Vec<String> = lines.iter()
                .map(|l| format!("\"{}\"", l.replace('\\', "\\\\").replace('"', "\\\"")))
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

    let nc: Vec<String> = no_coords.iter()
        .map(|n| format!("\"{}\"", n.replace('\\', "\\\\").replace('"', "\\\"")))
        .collect();

    let json = format!(
        r#"{{"file":"{}","orders":{},"customers":{},"matched":{},"products":{},"no_coords":[{}],"path":"{}"}}"#,
        fname, orders, customers, matched, products, nc.join(","),
        output_path.replace('\\', "\\\\")
    );

    progress(100, "执行完毕");
    Ok(json)
}

// ========== 明细处理 ==========

fn run_detail() -> Result<(), String> {
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
    for f in &files {
        process_single_detail(f)?;
    }
    Ok(())
}

fn process_single_detail(input: &str) -> Result<(), String> {
    let mut wb = open_workbook_auto(input).map_err(|e| format!("{}", e))?;
    let range = wb.worksheet_range_at(0).ok_or("no sheet")?.map_err(|e| format!("{}", e))?;
    let (rows, cols) = range.get_size();
    if rows < 3 { return Ok(()); }

    let mut ci = HashMap::new();
    for c in 0..cols {
        if let Some(Data::String(s)) = range.get_value((1, c as u32)) {
            for kw in ["商品名称", "规格", "单位", "实际价格", "订货数量", "金额", "备注"] {
                if s.contains(kw) && !ci.contains_key(kw) { ci.insert(kw, c); break; }
            }
        }
    }
    if !ci.contains_key("备注") && cols > 0 { ci.insert("备注", cols - 1); }
    let get = |k: &str| *ci.get(k).unwrap_or(&0);

    let mut remarks = Vec::new();
    for r in 2..rows {
        let name = cell_str(&range, r, get("商品名称"));
        let spec = cell_str(&range, r, get("规格"));
        let unit = cell_str(&range, r, get("单位"));
        let price = fmt_price(&cell_str(&range, r, get("实际价格")));
        let qty = fmt_qty(&cell_str(&range, r, get("订货数量")));
        let amount = fmt_amount(&cell_str(&range, r, get("金额")));
        remarks.push(if name.contains("赠品") {
            format!("{} {} {}：{}；", name, spec, unit, qty)
        } else {
            format!("{} {} {}：{}*{}= {}；", name, spec, unit, price, qty, amount)
        });
    }

    let nums: String = std::path::Path::new(input).file_name().unwrap().to_string_lossy()
        .chars().filter(|c| c.is_ascii_digit()).collect();
    write_xlsx(&format!("{}\\{}.xlsx", OUTPUT_DETAIL, nums), &range, rows, cols, get("备注"), &remarks)
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

fn run_order() -> Result<(Vec<String>, String, String, usize, usize, usize, i64), String> {
    let sales = r"D:\订单表格\销售订单.xls";
    let coords = r"D:\订单表格\客户经纬度.xls";
    let tmpl = r"D:\订单表格\优路达导入模板.xlsx";
    let detail = r"D:\明细表格\MXBG2";
    let out = r"D:\订单表格";

    for (p, n) in &[(sales, "销售订单"), (tmpl, "模板")] {
        if !std::path::Path::new(p).exists() { return Err(format!("缺少: {}", n)); }
    }

    progress(42, "步骤1: 读取订单...");
    let mut rows = read_sales(sales)?;
    if rows.is_empty() { return Err("无数据".into()); }

    progress(45, "步骤2: 经纬度匹配...");
    let unmatched = match_coords(coords, &mut rows);

    progress(50, "步骤3: 明细匹配...");
    let (_dt, matched) = match_details(detail, &mut rows);

    progress(55, "步骤4: 合并客户...");
    let orders = rows.len();
    rows = merge(rows);
    let customers = rows.len();

    let no_coords: Vec<String> = rows.iter()
        .filter(|r| r.coords.is_empty() && !r.customer_name.is_empty())
        .map(|r| r.location_name.clone()).collect();

    progress(60, "步骤5-6: 处理备注...");
    process_remarks(&mut rows);
    add_customer_info(&mut rows);

    progress(68, "步骤7: 生成模板...");
    let headers = read_headers(tmpl)?;

    progress(70, "步骤8: 保存...");
    let opath = save_output(&rows, &headers, out)?;
    let fname = std::path::Path::new(&opath).file_name()
        .map(|n| n.to_string_lossy().to_string()).unwrap_or_default();

    let products = count_products(detail);
    Ok((no_coords, fname, opath, orders, customers, matched, products))
}

// ========== 统计 ==========

fn compute_stats() -> (Vec<String>, i64) {
    let mut s: HashMap<(String, String, String), i64> = HashMap::new();
    if let Ok(e) = std::fs::read_dir(INPUT_DETAIL) {
        for en in e.flatten() {
            let n = en.file_name().to_string_lossy().to_string();
            if !n.ends_with(".xls") && !n.ends_with(".xlsx") || n.starts_with("~$") { continue; }
            let p = en.path().to_string_lossy().to_string();
            if let Ok(mut wb) = open_workbook_auto(&p) {
                if let Some(Ok(range)) = wb.worksheet_range_at(0) {
                    let (rows, cols) = range.get_size();
                    if rows < 3 { continue; }
                    let mut cn = None; let mut cs = None; let mut cu = None; let mut cq = None;
                    for c in 0..cols {
                        let cell = cell_str(&range, 1, c);
                        if cell.contains("商品名称") { cn = Some(c); }
                        else if cell.contains("规格") { cs = Some(c); }
                        else if cell.contains("单位") { cu = Some(c); }
                        else if cell.contains("订货数量") { cq = Some(c); }
                    }
                    let (a,b,c,d) = match (cn, cs, cu, cq) {
                        (Some(a),Some(b),Some(c),Some(d)) => (a,b,c,d), _ => continue,
                    };
                    for r in 2..rows {
                        let name = cell_str(&range, r, a);
                        if name.is_empty() { continue; }
                        if let Ok(q) = cell_str(&range, r, d).parse::<f64>() {
                            *s.entry((name, cell_str(&range, r, b), cell_str(&range, r, c))).or_default() += q as i64;
                        }
                    }
                }
            }
        }
    }
    if s.is_empty() { return (vec![], 0); }
    let mut items: Vec<_> = s.into_iter().collect();
    items.sort_by(|a, b| a.0.2.cmp(&b.0.2).then(a.0.0.cmp(&b.0.0)));
    let total: i64 = items.iter().map(|(_, v)| *v).sum();
    let (wn, ws, wu) = (46usize, 36usize, 10usize);
    let mut lines = Vec::new();
    let mut prev = None;
    for ((name, spec, unit), qty) in &items {
        if let Some(p) = prev { if p != unit { lines.push(String::new()); } }
        prev = Some(unit);
        lines.push(format!("  {}{}{}{:>6}", pad(name, wn), pad(spec, ws), pad(unit, wu), qty));
    }
    (lines, total)
}

fn count_products(dir: &str) -> i64 {
    let mut t = 0i64;
    if let Ok(e) = std::fs::read_dir(dir) {
        for en in e.flatten() {
            let n = en.file_name().to_string_lossy().to_string();
            if !n.ends_with(".xls") && !n.ends_with(".xlsx") || n.starts_with("~$") { continue; }
            if let Ok(mut wb) = open_workbook_auto(&en.path().to_string_lossy().to_string()) {
                if let Some(Ok(range)) = wb.worksheet_range_at(0) {
                    let (rows, cols) = range.get_size();
                    if rows < 3 { continue; }
                    if let Some(qc) = (0..cols).find(|&c| cell_str(&range, 1, c).contains("订货数量")) {
                        for r in 2..rows {
                            if let Ok(v) = cell_str(&range, r, qc).parse::<f64>() { t += v as i64; }
                        }
                    }
                }
            }
        }
    }
    t
}

// ========== 工具函数 ==========

fn cell_str(range: &calamine::Range<Data>, r: usize, c: usize) -> String {
    match range.get_value((r as u32, c as u32)) {
        Some(Data::String(s)) => s.trim().to_string(),
        Some(Data::Float(f)) => { let s = format!("{:.10}", *f); s.trim_end_matches('0').trim_end_matches('.').to_string() }
        Some(Data::Int(i)) => i.to_string(),
        Some(Data::Bool(b)) => b.to_string(),
        _ => String::new(),
    }
}

fn dw(s: &str) -> usize { s.chars().fold(0, |w, c| w + if c as u32 > 0x2E80 { 4 } else { 2 }) }
fn pad(s: &str, w: usize) -> String { format!("{}{}", s, " ".repeat(w.saturating_sub(dw(s)))) }

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
struct Row { customer_name:String, address:String, order_no:String, phone:String, contact_person:String, remark:String, location_name:String, coords:String, settle_status:String, unpaid_amount:f64 }

/// 表头单元格是否为该字段的精确列名（第一趟匹配用）
fn is_exact_col(field: &str, s: &str) -> bool {
    match field {
        "cu" => s == "客户单位",
        "ad" => s == "地址",
        "ct" => s == "联系信息",
        "on" => s == "订单号" || s == "销售单号",
        "rm" => s == "备注",
        "ss" => s == "结算状态" || s == "结账状态",
        _ => false,
    }
}

fn read_sales(path: &str) -> Result<Vec<Row>, String> {
    let mut wb = open_workbook_auto(path).map_err(|e| format!("{}", e))?;
    let range = wb.worksheet_range_at(0).ok_or("no sheet")?.map_err(|e| format!("{}", e))?;
    let (rows, cols) = range.get_size();
    let hr = (0..rows.min(5)).find(|&r| {
        (0..cols).filter(|&c| {
            let s = cell_str(&range, r, c);
            ["客户单位","地址","订单号","联系信息","备注"].iter().any(|k| s.contains(k))
        }).count() >= 3
    }).unwrap_or(0);

    // 列识别: 两趟匹配 + 已命中的字段不再被后面的列覆盖
    //   第1趟只认精确列名（"订单号" 必须完全等于 "订单号"）：新版销售订单在订单号后新增了
    //   「外部订单号」列，它 contains("订单号")，曾把真正的订单号列覆盖掉 → 订单号全空
    //   → 明细匹配 0 单、输出模板订单号列空白
    //   第2趟用 contains 兜底兼容带前后缀的表头，同时排除「外部订单号」等干扰列
    let mut map: HashMap<&str, usize> = HashMap::new();
    for pass in 0..2u8 {
        for c in 0..cols {
            let s = cell_str(&range, hr, c);
            let hit = if s.contains("客户单位") { Some("cu") }
                else if s.contains("地址") { Some("ad") }
                else if s.contains("联系信息") { Some("ct") }
                else if (s.contains("销售单号") || s.contains("订单号")) && !s.contains("外部") { Some("on") }
                else if s.contains("备注") { Some("rm") }
                else if s.contains("结算状态") || s.contains("结账状态") { Some("ss") }
                else { None };
            let hit = match hit { Some(h) => h, None => continue };
            if map.contains_key(hit) { continue; }
            if pass == 0 && !is_exact_col(hit, &s) { continue; }
            map.insert(hit, c);
        }
    }
    // 结算状态列未命中则回退到 S/U 列（旧版 S=18，新版 U=20，取大者）
    if !map.contains_key("ss") && cols > 20 { map.insert("ss", 20); }
    else if !map.contains_key("ss") && cols > 18 { map.insert("ss", 18); }
    if !map.contains_key("on") {
        for c in 0..cols {
            let s = cell_str(&range, hr, c);
            if s.contains("单号") && !s.contains("外部") { map.insert("on", c); break; }
        }
    }
    let g = |k: &str, r: usize| map.get(k).map(|&c| cell_str(&range, r, c)).unwrap_or_default();

    let mut result = Vec::new();
    for r in (hr+1)..rows {
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
        result.push(Row { customer_name:cu, address:ad, order_no:on, phone:ph, contact_person:cn, remark:rm2, location_name:loc, coords:String::new(), settle_status:g("ss", r), unpaid_amount:0.0 });
    }
    Ok(result)
}

fn match_coords(path: &str, rows: &mut [Row]) -> Vec<String> {
    let mut wb = match open_workbook_auto(path) { Ok(w) => w, Err(_) => return vec![] };
    let range = match wb.worksheet_range_at(0) { Some(Ok(r)) => r, _ => return vec![] };
    let (rn, cols) = range.get_size();
    let mut cc = None; let mut dc = None;
    for c in 0..cols {
        let s = cell_str(&range, 0, c);
        if s.contains("客户单位") { cc = Some(c); } else if s.contains("地图经纬度") { dc = Some(c); }
    }
    let (a, b) = match (cc, dc) { (Some(a), Some(b)) => (a, b), _ => return vec![] };
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

fn match_details(dir: &str, rows: &mut [Row]) -> (usize, usize) {
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

fn merge(rows: Vec<Row>) -> Vec<Row> {
    let mut g: HashMap<String, Vec<Row>> = HashMap::new();
    for r in rows { g.entry(r.customer_name.clone()).or_default().push(r); }
    g.into_values().map(|grp| {
        if grp.len() == 1 {
            let mut r = grp.into_iter().next().unwrap();
            if r.settle_status.contains("未结") {
                for l in r.remark.lines() {
                    if l.contains('=') && l.contains('；') {
                        if let Some(p) = l.rfind('=') {
                            let amt = l[p+1..].trim().trim_end_matches('；').replace(',',"").replace(' ',"");
                            if let Ok(a) = amt.parse::<f64>() { r.unpaid_amount += a; }
                        }
                    }
                }
            }
            return r;
        }
        let f = &grp[0];
        let mut os = Vec::new(); let mut rs: Vec<(String, bool)> = Vec::new();
        let mut any_unpaid = false; let mut any_paid = false;
        for r in &grp {
            if !r.order_no.is_empty() && !os.contains(&r.order_no) { os.push(r.order_no.clone()); }
            if !r.remark.is_empty() {
                let is_unpaid = r.settle_status.contains("未结");
                rs.push((r.remark.clone(), is_unpaid));
                if is_unpaid { any_unpaid = true; }
                else if r.settle_status.contains("已结") { any_paid = true; }
            }
        }
        // 合并规则: 全部未结→未结; 部分未结→部分; 全部已结→已结; 否则空
        let ss = if any_unpaid && any_paid { "部分未结".to_string() }
                 else if any_unpaid { "未结".to_string() }
                 else if any_paid { "已结".to_string() }
                 else { String::new() };
        let mixed = any_unpaid && any_paid;
        let mut unpaid_total = 0.0_f64;
        let parts: Vec<String> = rs.into_iter().map(|(remark, is_unpaid)| {
            if is_unpaid {
                // 累计未结订单的明细金额
                for l in remark.lines() {
                    if l.contains('=') && l.contains('；') {
                        if let Some(p) = l.rfind('=') {
                            let amt = l[p+1..].trim().trim_end_matches('；').replace(',',"").replace(' ',"");
                            if let Ok(a) = amt.parse::<f64>() { unpaid_total += a; }
                        }
                    }
                }
            }
            if mixed && is_unpaid {
                // 部分未结: 对未结订单的明细行末尾加 ♣
                remark.lines().map(|l| {
                    if l.contains('=') && l.contains('；') {
                        format!("{} ♣", l)
                    } else {
                        l.to_string()
                    }
                }).collect::<Vec<_>>().join("\n")
            } else {
                remark
            }
        }).collect();
        let mut rm = parts.join("\n");
        if !rm.is_empty() { rm.push_str(&format!("\n合并订单：{}", grp.len())); }
        Row { customer_name:f.customer_name.clone(), address:f.address.clone(), phone:f.phone.clone(),
              contact_person:f.contact_person.clone(), location_name:f.location_name.clone(),
              coords:f.coords.clone(), order_no:os.join(","), remark:rm, settle_status:ss, unpaid_amount:unpaid_total }
    }).collect()
}

fn process_remarks(rows: &mut [Row]) {
    for row in rows.iter_mut() {
        if row.remark.is_empty() { continue; }
        let lines: Vec<&str> = row.remark.lines().collect();
        let (mut t3, mut rm, mut mg, mut ot) = (vec![], vec![], vec![], vec![]);
        for l in &lines {
            if l.contains('=') && l.contains('；') { t3.push(l.to_string()); }
            else if l.starts_with("备注：") { rm.push(l.to_string()); }
            else if l.starts_with("合并订单：") { mg.push(l.to_string()); }
            else if !l.trim().is_empty() && l.trim() != &"-".repeat(10) { ot.push(l.to_string()); }
        }
        let mut total = 0.0;
        for l in &t3 { if let Some(p) = l.rfind('=') {
            if let Ok(a) = l[p+1..].trim().trim_end_matches('♣').trim().trim_end_matches('；').replace(',',"").replace(' ',"").parse::<f64>() { total += a; }
        }}
        // 未结订单: 将总金额回填，供输出模板"订单状态"列使用
        if row.settle_status.contains("未结") && row.unpaid_amount == 0.0 {
            row.unpaid_amount = total;
        }
        let num = if rm.len() > 1 {
            rm.iter().enumerate().map(|(i, l)| if i == 0 { format!("备注：{}、{}", i+1, &l[3..]) }
            else { format!("      {}、{}", i+1, &l[3..]) }).collect()
        } else { rm };
        let mut parts = vec![];
        if !t3.is_empty() {
            parts.extend(t3); parts.push(String::new());
            // 结算状态标记: 未结→" ♣♣", 部分未结→" ♣", 已结→" .", 其他/异常→" ?"
            let mark = if row.settle_status.contains("部分未结") { " ♣" }
                       else if row.settle_status.contains("未结") { " ♣♣" }
                       else if row.settle_status.contains("已结") { " ." }
                       else { " ?" };
            parts.push(format!("总金额：{:.2}元{}", total, mark));
        }
        if !mg.is_empty() { parts.extend(mg); }
        if !ot.is_empty() { parts.push(String::new()); parts.extend(ot); }
        if !num.is_empty() { parts.push(String::new()); parts.extend(num); }
        row.remark = parts.join("\n");
    }
}

fn add_customer_info(rows: &mut [Row]) {
    let now = chrono::Local::now();
    let target = if now.hour() >= 12 { now.date_naive().succ_opt().unwrap_or(now.date_naive()) }
    else { now.date_naive() };
    let ds = target.format("%Y年%m月%d日").to_string();
    for row in rows.iter_mut() {
        let parts: Vec<&str> = [&row.customer_name, &row.phone, &row.address, &row.contact_person]
            .iter().filter_map(|s| if s.is_empty() { None } else { Some(s.as_str()) }).collect();
        let info = parts.join("，");
        let mut p = vec![];
        if !info.is_empty() { p.push(info); p.push(String::new()); }
        if !row.remark.is_empty() { p.push(row.remark.clone()); }
        p.push(String::new()); p.push(format!("聚火配送：{}", ds));
        row.remark = p.join("\n");
    }
}

fn read_headers(path: &str) -> Result<Vec<String>, String> {
    let mut wb = open_workbook_auto(path).map_err(|e| format!("{}", e))?;
    let range = wb.worksheet_range_at(0).ok_or("no sheet")?.map_err(|e| format!("{}", e))?;
    Ok((0..range.get_size().1).map(|c| cell_str(&range, 0, c)).collect())
}

fn save_output(rows: &[Row], headers: &[String], dir: &str) -> Result<String, String> {
    use rand::Rng;
    let now = chrono::Local::now();
    let target = if now.hour() >= 12 { now.date_naive().succ_opt().unwrap_or(now.date_naive()) }
    else { now.date_naive() };
    let month = target.format("%m").to_string(); let day = target.format("%d").to_string();
    let prefix = format!("优路达导入模板{}{}", month, day);
    let base = rand::thread_rng().gen_range(1000..=9999);
    let mut me = 0u32;
    if let Ok(e) = std::fs::read_dir(dir) {
        for en in e.flatten() {
            let n = en.file_name().to_string_lossy().to_string();
            if n.starts_with(&prefix) && n.ends_with(".xlsx") {
                if let Some(s) = n.strip_prefix(&format!("{}-", &prefix)).and_then(|s| s.strip_suffix(".xlsx")) {
                    if let Ok(v) = s.parse::<u32>() { me = me.max(v); }
                }
            }
        }
    }
    let suffix = if base > me { base } else { let n = me + 1; if n > 9999 { 1000 + (n % 9000) } else { n } };
    let fp = format!("{}\\{}-{}.xlsx", dir, prefix, suffix);

    let mut wb = rust_xlsxwriter::Workbook::new(); let sh = wb.add_worksheet();
    for (c, h) in headers.iter().enumerate() { let _ = sh.write_string(0, c as u16, h); }
    let fields = ["客户名称","详细地址","订单号","手机号","客户经理","备注信息","地点名称","精准坐标","订单状态"];
    let col_map: HashMap<&str, usize> = fields.iter().filter_map(|&f| {
        headers.iter().position(|h| h.contains(f)).map(|c| (f, c))
    }).collect();

    for (r, row) in rows.iter().enumerate() {
        let rr = (r + 1) as u32;
        for (&f, &c) in &col_map {
            let v: String = match f {
                "客户名称" => row.customer_name.clone(),
                "详细地址" => row.address.clone(),
                "订单号" => row.order_no.clone(),
                "手机号" => row.phone.clone(),
                "客户经理" => row.contact_person.clone(),
                "备注信息" => row.remark.clone(),
                "地点名称" => row.location_name.clone(),
                "精准坐标" => row.coords.clone(),
                "订单状态" => {
                    if row.settle_status.contains("部分未结") {
                        format!("未结账-部分 ￥{:.2}", row.unpaid_amount)
                    } else if row.settle_status.contains("未结") {
                        format!("未结账-全部 ￥{:.2}", row.unpaid_amount)
                    } else if row.settle_status.contains("已结") {
                        "已结账".to_string()
                    } else {
                        "异常，请复核".to_string()
                    }
                }
                _ => continue,
            };
            if !v.is_empty() { let _ = sh.write_string(rr, c as u16, &v); }
        }
    }
    wb.save(&fp).map_err(|e| format!("{}", e))?;
    Ok(fp)
}
