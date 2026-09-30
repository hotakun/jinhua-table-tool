// 备注文本重组 + 模板写出 + 商品统计
// v3.1.3 从 engine.rs 拆出（"输出侧"逻辑集中在这里）
use std::collections::HashMap;
use calamine::{open_workbook_auto, Reader};
use chrono::Timelike;
use crate::{cell_str, dw, has_error, spec_hit, Diag, Row, DETAIL_FIELDS};

pub(crate) fn merge(rows: Vec<Row>) -> Vec<Row> {
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
                    if l.contains('=') && l.contains('；') { format!("{} ♣", l) } else { l.to_string() }
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

pub(crate) fn process_remarks(rows: &mut [Row]) {
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

pub(crate) fn add_customer_info(rows: &mut [Row]) {
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

pub(crate) fn read_headers(path: &str) -> Result<Vec<String>, String> {
    let mut wb = open_workbook_auto(path).map_err(|e| format!("{}", e))?;
    let range = wb.worksheet_range_at(0).ok_or("no sheet")?.map_err(|e| format!("{}", e))?;
    Ok((0..range.get_size().1).map(|c| cell_str(&range, 0, c)).collect())
}

pub(crate) fn save_output(rows: &[Row], headers: &[String], dir: &str,
                          diags: &mut Vec<Diag>, force: bool) -> Result<String, String> {
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

    // 模板列识别：精确名优先 → 别名兜底（写错列的风险从根上避免）
    let tfields: [(&str, &[&str], bool); 9] = [
        ("客户名称", &["客户名称", "客户"], true),
        ("详细地址", &["详细地址", "地址"], false),
        ("订单号", &["订单号", "订单编号"], false),
        ("手机号", &["手机号", "联系电话", "电话"], false),
        ("客户经理", &["客户经理", "业务员", "负责人"], false),
        ("备注信息", &["备注信息", "备注"], true),
        ("地点名称", &["地点名称", "地点"], false),
        ("精准坐标", &["精准坐标", "经纬度", "坐标"], false),
        ("订单状态", &["订单状态", "状态"], false),
    ];
    let mut col_map: HashMap<&str, usize> = HashMap::new();
    for (f, cands, required) in tfields.iter() {
        let mut hit: Option<usize> = None;
        for pass in 0..2u8 {
            if hit.is_some() { break; }
            for (i, h) in headers.iter().enumerate() {
                let ok = if pass == 0 { cands.iter().any(|n| h == n) }
                         else { cands.iter().any(|n| h.contains(n)) };
                if ok { hit = Some(i); break; }
            }
        }
        match hit {
            Some(i) => { col_map.insert(f, i); }
            None => {
                if *required {
                    diags.push(Diag::error(f, &format!("模板缺少必需列「{}」", f),
                        format!("模板表头: {}", headers.join(" / "))));
                } else {
                    diags.push(Diag::warn(f, &format!("模板未找到「{}」列，该列不写入", f),
                        format!("模板表头: {}", headers.join(" / "))));
                }
            }
        }
    }
    if has_error(diags) && !force {
        return Err("模板缺少必需列，已中止（可用「仍然生成」继续）".into());
    }
    if !col_map.contains_key("订单状态") {
        diags.push(Diag::warn("订单状态", "模板没有「订单状态」列，结算状态不会写入表格",
            "备注文本里的 ♣ / . 标记仍然有效".to_string()));
    }

    let mut wb = rust_xlsxwriter::Workbook::new(); let sh = wb.add_worksheet();
    for (c, h) in headers.iter().enumerate() { let _ = sh.write_string(0, c as u16, h); }

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

pub(crate) fn compute_stats() -> (Vec<String>, i64) {
    let mut s: HashMap<(String, String, String), i64> = HashMap::new();
    if let Ok(e) = std::fs::read_dir(crate::INPUT_DETAIL) {
        for en in e.flatten() {
            let n = en.file_name().to_string_lossy().to_string();
            if !n.ends_with(".xls") && !n.ends_with(".xlsx") || n.starts_with("~$") { continue; }
            let p = en.path().to_string_lossy().to_string();
            if let Ok(mut wb) = open_workbook_auto(&p) {
                if let Some(Ok(range)) = wb.worksheet_range_at(0) {
                    let (rows, cols) = range.get_size();
                    if rows < 3 { continue; }
                    let mut ci: HashMap<&'static str, usize> = HashMap::new();
                    for pass in 0..2u8 {
                        for c in 0..cols {
                            let cell = cell_str(&range, 1, c);
                            for sp in DETAIL_FIELDS {
                                if ci.contains_key(sp.key) { continue; }
                                if spec_hit(sp, &cell, pass == 0) { ci.insert(sp.key, c); }
                            }
                        }
                    }
                    let (a, b, c, d) = match (ci.get("name"), ci.get("spec"), ci.get("unit"), ci.get("qty")) {
                        (Some(&a), Some(&b), Some(&c), Some(&d)) => (a, b, c, d),
                        _ => continue,
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
        lines.push(format!("  {}{}{}{:>6}", pad_(name, wn), pad_(spec, ws), pad_(unit, wu), qty));
    }
    (lines, total)
}

/// 等宽填充：内容超宽时按显示宽度截断并补 "…"，保证后面的列不被顶歪
/// （例如规格 '570/580/680/850/960特厚大*1800个' 显示宽度 64 > 36，
///   以前只补空格不截断，会把「单位」「数量」整体往右顶，看起来像串列）
fn pad_(s: &str, w: usize) -> String {
    let cur = dw(s);
    if cur <= w {
        return format!("{}{}", s, " ".repeat(w - cur));
    }
    let limit = w.saturating_sub(2);        // 给 "…"（显示宽度 2）留位置
    let mut out = String::new();
    let mut used = 0usize;
    for ch in s.chars() {
        let cw = if ch as u32 > 0x2E80 { 4 } else { 2 };
        if used + cw > limit { break; }
        out.push(ch);
        used += cw;
    }
    out.push('…');
    used += 2;
    if used < w { out.push_str(&" ".repeat(w - used)); }
    out
}

pub(crate) fn count_products(dir: &str) -> i64 {
    let mut t = 0i64;
    if let Ok(e) = std::fs::read_dir(dir) {
        for en in e.flatten() {
            let n = en.file_name().to_string_lossy().to_string();
            if !n.ends_with(".xls") && !n.ends_with(".xlsx") || n.starts_with("~$") { continue; }
            if let Ok(mut wb) = open_workbook_auto(&en.path().to_string_lossy().to_string()) {
                if let Some(Ok(range)) = wb.worksheet_range_at(0) {
                    let (rows, cols) = range.get_size();
                    if rows < 3 { continue; }
                    let qc = (0..cols).find(|&c| {
                        let s = cell_str(&range, 1, c);
                        spec_hit(&DETAIL_FIELDS[4], &s, true) || spec_hit(&DETAIL_FIELDS[4], &s, false)
                    });
                    if let Some(qc) = qc {
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
