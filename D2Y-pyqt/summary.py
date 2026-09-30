#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""商品汇总（分类 > 商品 × 客户）—— 生成手机友好的单文件 HTML。

被 main.py 调用：程序执行时先在明细被清空前 collect() 聚合数据，
用户点「打开并退出」时再用 render() 落盘 HTML（与 TXT 同时生成）。

也可以单独运行（手动生成一份）：
    cd /d D:\WFR\D2Y-pyqt
    python3 summary.py

数据来源（只读，不会改动任何原始文件）：
    D:\\明细表格\\销售订单详细 - <订单号>.xls   分类/商品名称/规格/单位/订货数量
    D:\\订单表格\\销售订单.xls                  订单号 → 客户单位（按订单号 join）

输出：
    D:\\订单表格\\每日详情\\商品汇总MMDD.html
"""
import os
import re
import sys
from collections import defaultdict
from datetime import datetime

DET_DIR = r"D:\明细表格"
SALES = r"D:\订单表格\销售订单.xls"
OUT_DIR = r"D:\订单表格\每日详情"

# 大类显示顺序（按指定优先级；不在表里的大类排最后）
CAT_ORDER = ["打包盒", "纸碗", "饭盒", "调料盒", "纸", "杯", "袋",
             "勺", "筷", "签", "洗护清洁"]

# 大类配色（淡色，按上面顺序循环取用）
PALETTE = ["#eef6ff", "#eefaf1", "#fff8e8", "#fdeef5", "#f1effc", "#e9f7f8",
           "#f2f7ea", "#fdf0e8"]

CSS = """
* { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
body { margin:0; padding:0 10px 24px; background:#f5f6f8; color:#222;
       font:16px/1.45 -apple-system,"Microsoft YaHei","PingFang SC",sans-serif; }
/* 标题区在吸顶栏之外，随内容自然滚走；只有搜索栏吸顶 ——
   纯 CSS 实现，不监听 scroll，因此不会与手机浏览器地址栏伸缩互相干扰 */
.top { padding:10px 0 0; }
header { position:sticky; top:0; z-index:5; background:#f5f6f8;
         padding:0 0 8px; border-bottom:1px solid transparent; }
h1 { font-size:18px; margin:0 0 2px; }
.stat { font-size:12px; color:#666; margin-bottom:8px; }
.bar { display:flex; gap:6px; align-items:stretch; }
#q { flex:1 1 auto; min-width:0; padding:9px 12px; font-size:15px;
     border:1px solid #d0d3d8; border-radius:8px; background:#fff;
     color:inherit; outline:none; }
#tgl { flex:0 0 auto; padding:9px 12px; font-size:14px; white-space:nowrap;
       border:1px solid #d0d3d8; border-radius:8px; background:#fff; color:inherit; }
details { border-radius:8px; margin-bottom:6px; overflow:hidden;
          box-shadow:0 1px 2px rgba(0,0,0,.05); }
summary { display:flex; align-items:center; gap:4px; padding:8px 10px;
          cursor:pointer; list-style:none; line-height:1.3; }
summary::-webkit-details-marker { display:none; }
summary::before { content:"\\25B8"; color:#9aa0a6; font-size:10px;
                  flex:0 0 10px; margin-left:-2px; }
details[open] summary::before { content:"\\25BE"; }
.nm { flex:0 0 auto; font-size:14px; font-weight:600; white-space:nowrap; }
.sub { flex:0 1 auto; min-width:0; margin-left:auto; text-align:right;
       font-size:12px; color:#7c838a;
       white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.tq { flex:0 0 auto; min-width:46px; text-align:right; font-size:16px;
      font-weight:700; color:#e6791a; }
ul { margin:0; padding:0 0 4px; list-style:none; background:#fff; }
li { display:flex; gap:8px; padding:7px 10px 7px 26px; font-size:13.5px;
     color:#444; border-top:1px solid #eef0f2; }
.cn { flex:1 1 auto; min-width:0; white-space:nowrap; overflow:hidden;
      text-overflow:ellipsis; }
.cq { flex:0 0 auto; min-width:44px; text-align:right; color:#333; }
/* 搜索命中的店铺行：只把文字变蓝（不加粗、不变大、不改底色） */
li.bhit .cn, li.bhit .cq { color:#0b5fbe; }
footer { margin-top:10px; font-size:12px; color:#9aa0a6; text-align:center; }
"""

JS = """
const q = document.getElementById('q');
const list = document.getElementById('list');
const tgl = document.getElementById('tgl');
const items = () => list.querySelectorAll('details');
const vis = () => Array.from(items()).filter(d => d.style.display !== 'none');

q.addEventListener('input', () => {
  const raw = q.value.trim();
  const k = raw.toLowerCase();
  /* 纯字母数字 → 客户代码精确匹配；含文字 → 模糊匹配 */
  const codeMode = !!raw && /^[A-Za-z0-9]+$/.test(raw);
  items().forEach(d => {
    /* 商品名/规格/分类只做模糊匹配；纯代码模式下不参与，免得带出一堆相近商品 */
    const selfHit = !!k && !codeMode && (d.dataset.text || '').includes(k);
    let custHit = false;
    d.querySelectorAll('li').forEach(li => {
      const t = (li.dataset.c || '').toLowerCase();
      const code = (t.split(/[\s★\-_]+/)[0] || '');
      const h = !!k && (codeMode ? code === k : t.includes(k));
      li.classList.toggle('bhit', h);      /* 命中的店铺行标蓝 */
      if (h) custHit = true;
    });
    const show = !k || selfHit || custHit;
    d.style.display = show ? '' : 'none';
    d.open = !!k && show;                  /* 搜索时自动展开命中的商品 */
  });
  syncBtn();     /* 搜索后同步按钮文字：结果都展开时应显示"全部收起" */
});

/* 按当前可见项的开合状态刷新按钮文字 */
function syncBtn() {
  const anyClosed = vis().some(d => !d.open);
  tgl.textContent = anyClosed ? '全部展开' : '全部收起';
}

tgl.addEventListener('click', () => {
  const anyClosed = vis().some(d => !d.open);
  vis().forEach(d => { d.open = anyClosed; });
  syncBtn();
});

/* 说明：标题与总览放在吸顶栏之外，会随内容自然滚走；
   搜索栏用 position:sticky 吸顶，完全不用 JS 监听滚动，
   所以不会和手机浏览器地址栏的伸缩互相触发抖动 */
"""


def _log(quiet, msg):
    if not quiet:
        print(msg)


def cell(sh, r, c):
    if c is None or r >= sh.nrows or c >= sh.ncols:
        return ""
    v = sh.cell_value(r, c)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    return str(v).strip()


def find_col(sh, header_row, keywords, exclude=()):
    """按关键字找列（表头行内）；命中排除词的列跳过"""
    for c in range(sh.ncols):
        s = cell(sh, header_row, c)
        if any(k in s for k in keywords) and not any(x in s for x in exclude):
            return c
    return None


def split_cat(c):
    """'大类>小类' → (大类, 小类)；没有 > 就当大类"""
    s = (c or "").strip().replace("＞", ">").replace("／", "/")
    if not s:
        return ("未分类", "")
    if ">" in s:
        a, b = s.split(">", 1)
        return (a.strip() or "未分类", b.strip())
    return (s, "")


def cat_rank(big):
    """大类在 CAT_ORDER 中的次序；未列出的排到最后
       先精确匹配，再取"关键词被大类名包含"中的最长项"""
    b = (big or "").strip()
    for i, k in enumerate(CAT_ORDER):
        if b == k:
            return i
    best = None
    for i, k in enumerate(CAT_ORDER):
        if k and k in b:
            cand = (len(k), -i)
            if best is None or cand > best:
                best = cand
    if best is not None:
        return -best[1]
    for i, k in enumerate(CAT_ORDER):
        if b and b in k:
            return i
    return len(CAT_ORDER) + 100


def natkey(s):
    """自然排序键：字符串里的数字段按"数值"大小比
       （这样 200 排在 1000 前面，而不是按字符 '1'<'2' 把 1000 排前面）"""
    out = []
    for p in re.split(r"(\d+)", str(s)):
        if not p:
            continue
        out.append((1, int(p), "") if p.isdigit() else (0, 0, p))
    return tuple(out)


def num(v):
    return str(int(v)) if float(v).is_integer() else ("%.2f" % v)


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def load_customers(quiet=False):
    """订单号 → 客户单位"""
    try:
        import xlrd
    except Exception as e:
        _log(quiet, "!! 缺少 xlrd，无法读取 .xls：%s" % e)
        return {}
    m = {}
    try:
        bk = xlrd.open_workbook(SALES)
    except Exception as e:
        _log(quiet, "!! 打不开销售订单.xls：%s" % e)
        return m
    sh = bk.sheet_by_index(0)
    hr = 0
    c_on = find_col(sh, hr, ["订单号", "销售单号"], exclude=["外部", "物流"])
    c_cu = find_col(sh, hr, ["客户单位", "客户名称"])
    if c_on is None or c_cu is None:
        _log(quiet, "!! 销售订单表头没找到「订单号」或「客户单位」列。表头前 20 列：%s"
             % [cell(sh, hr, c) for c in range(min(sh.ncols, 20))])
        return m
    for r in range(hr + 1, sh.nrows):
        on = cell(sh, r, c_on)
        cu = cell(sh, r, c_cu)
        if on and cu:
            m.setdefault(on, cu)
    _log(quiet, "订单 → 客户 映射：%d 条" % len(m))
    return m


def collect(quiet=False):
    """只读聚合（不改动任何文件）。返回 dict；失败时 {"error": "原因"}。

       注意：要在 D:\\明细表格 被清空之前调用。
    """
    try:
        import xlrd
    except Exception as e:
        return {"error": "缺少 xlrd 库：%s" % e}

    cust = load_customers(quiet)

    if not os.path.isdir(DET_DIR):
        return {"error": "找不到目录 %s" % DET_DIR}
    files = [f for f in os.listdir(DET_DIR)
             if f.lower().endswith(".xls") and not f.startswith("~$")]
    if not files:
        return {"error": r"D:\明细表格 下没有 .xls 文件"}

    # (商品名称, 规格, 单位) -> {"cat": (大类, 小类), "cust": {客户: 数量}}
    agg = {}
    bad = 0
    miss_cust = 0
    for fn in files:
        m = re.search(r"(\d+)", fn)
        on = m.group(1) if m else ""
        if on and on in cust:
            cu = cust[on]
        else:
            cu = ("未知客户(%s)" % on) if on else "未知客户(无订单号)"
            miss_cust += 1
        try:
            bk = xlrd.open_workbook(os.path.join(DET_DIR, fn))
            sh = bk.sheet_by_index(0)
        except Exception:
            bad += 1
            continue
        if sh.nrows < 3:
            continue

        hr = 1   # 明细表头在第 2 行（与引擎一致）
        c_cat = find_col(sh, hr, ["分类"])
        c_name = find_col(sh, hr, ["商品名称"])
        c_spec = find_col(sh, hr, ["规格"])
        c_unit = find_col(sh, hr, ["单位"], exclude=["数量"])
        c_qty = find_col(sh, hr, ["订货数量", "数量"], exclude=["出库", "库存"])
        if c_name is None or c_qty is None:
            bad += 1
            continue

        for r in range(2, sh.nrows):
            name = cell(sh, r, c_name)
            if not name:
                continue
            try:
                qty = float(cell(sh, r, c_qty) or 0)
            except Exception:
                continue
            key = (name, cell(sh, r, c_spec), cell(sh, r, c_unit))
            e = agg.get(key)
            if e is None:
                e = {"cat": split_cat(cell(sh, r, c_cat) if c_cat is not None else ""),
                     "cust": defaultdict(float)}
                agg[key] = e
            e["cust"][cu] += qty

    if not agg:
        return {"error": "没有汇总到任何商品行"}

    # 排序：大类（按 CAT_ORDER 指定顺序）→ 小类 → 商品名称 → 规格 → 单位
    items = sorted(agg.items(),
                   key=lambda kv: (cat_rank(kv[1]["cat"][0]), natkey(kv[1]["cat"][1]),
                                   natkey(kv[0][0]), natkey(kv[0][1]), natkey(kv[0][2])))

    # 同一大类统一底色
    color_of = {}
    for _, v in items:
        b = v["cat"][0]
        if b not in color_of:
            color_of[b] = PALETTE[len(color_of) % len(PALETTE)]

    cust_set = set()
    for _, v in items:
        cust_set.update(v["cust"].keys())
    total_all = sum(sum(v["cust"].values()) for _, v in items)

    return {"items": items, "color_of": color_of, "files": len(files),
            "bad": bad, "miss": miss_cust, "cust_n": len(cust_set), "total": total_all}


def build_html(data):
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    items = data["items"]
    color_of = data["color_of"]
    cards = []
    for (name, spec, unit), v in items:
        big, small = v["cat"]
        cust = v["cust"]
        text = " ".join([big, small, name, spec, unit] + list(cust.keys())).lower()
        rows = "".join(
            '<li data-c="%s"><span class="cn">%s</span><span class="cq">%s</span></li>'
            % (esc(cu.lower()), esc(cu), num(q))
            for cu, q in sorted(cust.items(), key=lambda kv: (-kv[1], natkey(kv[0])))
        )
        sub = (spec + " " + unit).strip()
        cards.append(
            '<details data-text="%s" style="background:%s"><summary>'
            '<span class="nm">%s</span>%s<span class="tq">%s</span>'
            '</summary><ul>%s</ul></details>'
            % (esc(text), color_of.get(big, PALETTE[0]), esc(name),
               ('<span class="sub">%s</span>' % esc(sub)) if sub else "",
               num(sum(cust.values())), rows)
        )

    warn = ""
    if data["bad"] or data["miss"]:
        warn = ('<div class="stat" style="color:#b45309;">'
                '注意：%d 个明细文件读取失败，%d 个订单未匹配到客户</div>'
                % (data["bad"], data["miss"]))

    return """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>商品汇总 %s</title><style>%s</style></head><body>
<div class="top">
  <h1>商品汇总</h1>
  <div class="stat">商品 %d 种 · 客户 %d 家 · 总数量 %s · 明细文件 %d 个 · 生成于 %s</div>
  %s
</div>
<header>
  <div class="bar">
    <input id="q" placeholder="搜索商品 / 规格 / 分类 / 客户…">
    <button id="tgl">全部展开</button>
  </div>
</header>
<div id="list">%s</div>
<footer>点商品行可展开各客户数量</footer>
<script>%s</script></body></html>""" % (
        now, CSS, len(items), data["cust_n"], num(data["total"]),
        data["files"], now, warn, "".join(cards), JS)


def render(data, out_dir=None, quiet=False):
    """把 collect() 的结果落盘为 HTML，返回文件路径（失败返回 ""）。"""
    if not data or data.get("error"):
        _log(quiet, "!! 汇总数据不可用：%s" % (data or {}).get("error"))
        return ""
    out_dir = out_dir or OUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "商品汇总%s.html" % datetime.now().strftime("%m%d"))
    with open(out, "w", encoding="utf-8") as f:
        f.write(build_html(data))
    _log(quiet, "已生成：%s (%.0f KB)" % (out, os.path.getsize(out) / 1024.0))
    return out


def generate(out_dir=None, do_open=False, quiet=False):
    """一步到位：聚合 + 落盘（供手动/独立运行）。

       返回 {"path": ..., "items": n, ...}；失败返回 {"error": "..."}。
    """
    data = collect(quiet=quiet)
    if data.get("error"):
        return data
    path = render(data, out_dir=out_dir, quiet=quiet)
    if not path:
        return {"error": "写出 HTML 失败"}
    _log(quiet, "商品 %d 种 | 大类 %d 个 | 客户 %d 家 | 总数量 %s"
         % (len(data["items"]), len(data["color_of"]), data["cust_n"], num(data["total"])))
    _log(quiet, "明细文件 %d 个（读取失败 %d 个；未匹配到客户的 %d 个）"
         % (data["files"], data["bad"], data["miss"]))
    if do_open:
        try:
            os.startfile(path)
        except Exception:
            pass
    return {"path": path, "items": len(data["items"]), "bigs": len(data["color_of"]),
            "cust_n": data["cust_n"], "total": data["total"],
            "files": data["files"], "bad": data["bad"], "miss": data["miss"]}


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    r = generate(do_open=True)
    if r.get("error"):
        print("!! %s" % r["error"])
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
