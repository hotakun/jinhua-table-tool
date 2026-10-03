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
    // 地址列（可选）：同编号多店铺时**唯一可靠的区分依据**，没有就只能留空报警
    let ad_col = map.get("ad").copied();
    if ad_col.is_none() {
        diags.push(Diag::warn("客户经纬度", "客户经纬度表里没有地址列",
            "同一个客户编号对应多家店铺时将无法区分，这些客户会留空处理。建议给该表补一列「地址」".to_string()));
    }

    // ── 归一化与提取 ──────────────────────────────────────────────
    // 去空白 + 全角标点转半角 + 去掉尾部标点（如 "c45 ...(东塔路店)." 末尾那个句点）
    fn norm(s: &str) -> String {
        let t: String = s.chars()
            .filter(|c| !c.is_whitespace())
            .map(|c| match c {
                '（' => '(', '）' => ')', '，' => ',', '．' => '.', '－' => '-',
                '、' => ',', '：' => ':', '；' => ';',
                _ => c,
            })
            .collect();
        t.trim_end_matches(|c| ".。，,、；;：:!！?？·".contains(c)).to_string()
    }
    // 店名 = 去掉开头编号（字母数字）后的剩余部分。
    // 注意：客户单位里**可能没有空格**（a331童记北京烤鸭店(东城店)），不能按空格切！
    fn shop_of(name: &str) -> String {
        norm(name).trim_start_matches(|c: char| c.is_ascii_alphanumeric()).to_string()
    }
    // 提取客户单位中的代码（第一个空格/★前的字母数字）
    fn extract_code(name: &str) -> String {
        let prefix = name.split(|c: char| c == ' ' || c == '★').next().unwrap_or("");
        prefix.chars().filter(|c| c.is_alphanumeric()).collect::<String>().to_uppercase()
    }
    // 地址打分：地址只是**补充**判据，不许「取最相似」——必须明显胜出才认
    fn addr_score(a: &str, b: &str) -> u32 {
        let a = norm(a); let b = norm(b);
        if a.is_empty() || b.is_empty() { return 0; }
        if a == b { return 1000; }
        if a.contains(&b) || b.contains(&a) { return 500; }
        const SEP: &str = "省市区县镇乡村路街道巷弄号栋幢楼室园区广场大厦";
        let mut hit = 0u32;
        for s in a.split(|c: char| SEP.contains(c)) {
            if s.chars().count() >= 2 && b.contains(s) { hit += 1; }
        }
        hit * 10
    }
    // 在候选里按地址挑唯一明显胜出的，挑不出返回 None；idxs 是要参与比较的下标
    fn pick_by_addr(myaddr: &str, cands: &[(String, String, String)], idxs: &[usize]) -> Option<usize> {
        if idxs.is_empty() { return None; }
        let mut scored: Vec<(u32, usize)> = idxs.iter()
            .map(|&i| (addr_score(myaddr, &cands[i].1), i)).collect();
        scored.sort_by(|x, y| y.0.cmp(&x.0));
        let top = scored[0].0;
        let second = scored.get(1).map(|x| x.0).unwrap_or(0);
        if top >= 500 && top > second { Some(scored[0].1) } else { None }
    }

    // 建映射：编号 -> Vec<(归一化店名, 归一化地址, 坐标)>
    // 旧代码这里是 code -> 单个坐标，同编号的第二家店会被静默丢弃 → 送错货
    let mut code_map: HashMap<String, Vec<(String, String, String)>> = HashMap::new();
    for r in 1..rn {
        let n = cell_str(&range, r, a); let c = cell_str(&range, r, b);
        let ad = ad_col.map(|x| cell_str(&range, r, x)).unwrap_or_default();
        if !n.is_empty() && !c.is_empty() {
            let code = extract_code(&n);
            if !code.is_empty() {
                code_map.entry(code).or_default().push((shop_of(&n), norm(&ad), c));
            }
        }
    }

    // ── 匹配：①店名精确（多条时用地址定夺）→ ②编号唯一放松 → ③地址补充 → ④留空报警 ──
    let mut un = Vec::new();         // 未匹配（红色展示）
    let mut ambiguous = Vec::new();  // 定不了 → 留空
    let mut loose = Vec::new();      // 编号唯一但店名不一致 → 放行 + 提醒
    for row in rows.iter_mut() {
        let code = extract_code(&row.customer_name);
        let cands = match code_map.get(&code) {
            Some(v) => v,
            None => {
                if !row.customer_name.trim().is_empty() { un.push(row.customer_name.clone()); }
                continue;
            }
        };
        let me = shop_of(&row.customer_name);
        let myaddr = norm(&row.address);

        // ① 店名归一化后完全相等
        let exact: Vec<usize> = cands.iter().enumerate()
            .filter(|(_, (s, _, _))| *s == me).map(|(i, _)| i).collect();
        if exact.len() == 1 { row.coords = cands[exact[0]].2.clone(); continue; }

        // ①-b 命中多条：管理人员可能误录了同编号同店名的两行 → 用地址定夺
        if exact.len() > 1 {
            if let Some(i) = pick_by_addr(&myaddr, cands, &exact) {
                row.coords = cands[i].2.clone();
                continue;
            }
            un.push(row.customer_name.clone());
            ambiguous.push(format!("{}（表内同编号同店名有 {} 行，地址也分不出）",
                row.customer_name, exact.len()));
            continue;
        }

        // ② 编号唯一 → 放松放行（用户允许的宽松分支）
        if cands.len() == 1 {
            row.coords = cands[0].2.clone();
            if !me.is_empty() {
                loose.push(format!("{}（表内唯一候选店名：{}）", row.customer_name, cands[0].0));
            }
            continue;
        }

        // ③ 编号多店、店名对不上 → 地址补充，必须明显胜出
        let all: Vec<usize> = (0..cands.len()).collect();
        if let Some(i) = pick_by_addr(&myaddr, cands, &all) {
            row.coords = cands[i].2.clone();
            continue;
        }

        // ④ 都对不上 → 留空，绝不猜（宁可没有坐标，也不能送错）
        un.push(row.customer_name.clone());
        ambiguous.push(format!("{}（候选：{}）", row.customer_name,
            cands.iter().map(|(s, ad, _)| format!("{} | {}", s, ad))
                 .collect::<Vec<_>>().join(" / ")));
    }

    if !ambiguous.is_empty() {
        diags.push(Diag::warn("客户经纬度",
            &format!("有 {} 个客户编号对应多家店铺，店名与地址都对不上，已留空以免送错，请人工核对",
                     ambiguous.len()),
            ambiguous.join("；")));
    }
    if !loose.is_empty() {
        diags.push(Diag::warn("客户经纬度",
            &format!("有 {} 个客户编号在表里唯一但店名不一致，已按该唯一记录填坐标，请确认",
                     loose.len()),
            loose.join("；")));
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
