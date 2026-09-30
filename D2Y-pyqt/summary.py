#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""商品汇总（分类 > 商品 × 客户）—— 生成手机友好的单文件 HTML。

被 main.py 调用：程序执行时先在明细被清空前 collect() 聚合数据，
用户点「打开并退出」时再用 render() 落盘 HTML（与 TXT 同时生成）。

也可以单独运行（手动生成一份）：
    cd /d D:\\WFR\\D2Y-pyqt
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

# 指定大类下小类的显示顺序；未列入的小类按名称自然序排在后面
SUB_ORDER = {
    "打包盒": ["方盒", "圆碗", "圆盆", "美式", "格子盒", "异形盒"],
}

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
.bar { display:flex; gap:6px; align-items:stretch; flex-wrap:nowrap; }
/* 视图切换按钮：单字，尽量给搜索框腾地方（当前视图高亮） */
.vbtn { flex:0 0 auto; min-width:38px; padding:9px 8px; font-size:14px;
        text-align:center; white-space:nowrap;
        border:1px solid #d0d3d8; border-radius:8px; background:#fff; color:#555; }
.vbtn.on { background:#0b5fbe; border-color:#0b5fbe; color:#fff; font-weight:600; }
/* 吸顶栏保持“常驻吸顶”（不随滚动隐藏）；全程不监听 scroll，
   因此不会与手机浏览器地址栏的伸缩互相干扰 */
/* 按客户视图：客户是卡片头，商品行沿用同一套四段式样式 */
.cnm { flex:1 1 auto; min-width:0; font-size:14px; font-weight:600;
       white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
/* 客户视角：客户栏自带底色（行内 style），展开后的商品行一律纯白底、
   黑字、不加粗，也不上橙色；只有搜索命中的行才变蓝 */
#listC li { background:#fff; }
#listC li .nm { font-weight:400; }
#listC li .sub, #listC li .unit { color:#3a3f45; }
#listC li .tq { font-size:14px; font-weight:400; color:#222; }
#listC li.bhit .nm, #listC li.bhit .sub,
#listC li.bhit .unit, #listC li.bhit .tq { color:#0b5fbe; }
#listC[hidden] { display:none; }
#q { flex:1 1 auto; min-width:0; padding:9px 12px; font-size:15px;
     border:1px solid #d0d3d8; border-radius:8px; background:#fff;
     color:inherit; outline:none; }
#tgl { flex:0 0 auto; min-width:34px; padding:9px 8px; font-size:14px;
       text-align:center; white-space:nowrap;
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
/* 规格：占中间剩余空间，右对齐，空间不够才省略 */
.sub { flex:0 1 auto; min-width:0; margin-left:auto; text-align:right;
       font-size:12px; color:#7c838a;
       white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
/* 单位：独立窄列（固定宽度），永远不会被长规格挤占 */
.unit { flex:0 0 auto; min-width:2.4em; text-align:right; font-size:12px;
        color:#7c838a; white-space:nowrap; }
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
const tgl = document.getElementById('tgl');
const vP = document.getElementById('vP');
const vC = document.getElementById('vC');
const listP = document.getElementById('list');
const listC = document.getElementById('listC');
let view = 'p';        /* 每次打开都默认在「商」视图（不记忆上次选择）*/
const curList = () => (view === 'c' ? listC : listP);
const items = () => curList().querySelectorAll('details');
const vis = () => Array.from(items()).filter(d => d.style.display !== 'none');

/* 切换视图：只切 display（两份都是服务端渲染好的 DOM，切换零等待）*/
function applyView() {
  listP.hidden = (view !== 'p');
  listC.hidden = (view !== 'c');
  vP.classList.toggle('on', view === 'p');
  vC.classList.toggle('on', view === 'c');
  doSearch();        /* 换视图后按同一个关键词重新过滤，不会把搜索“弄没” */
}

function doSearch() {
  const raw = q.value.trim();
  const k = raw.toLowerCase();
  /* 纯字母数字 → 客户/店铺代码精确匹配；含文字 → 模糊匹配 */
  const codeMode = !!raw && /^[A-Za-z0-9]+$/.test(raw);
  items().forEach(d => {
    let selfHit = false;
    let subHit = false;
    if (view === 'p') {
      /* 商品视角：卡片头是商品，子行是客户 */
      selfHit = !!k && !codeMode && (d.dataset.text || '').includes(k);
      d.querySelectorAll('li').forEach(li => {
        const t = (li.dataset.c || '').toLowerCase();
        const code = (t.split(/[\\s★\\-_]+/)[0] || '');
        const h = !!k && (codeMode ? code === k : t.includes(k));
        li.classList.toggle('bhit', h);      /* 命中的店铺行标蓝 */
        if (h) subHit = true;
      });
    } else {
      /* 客户视角：卡片头是客户，子行是商品 */
      const code = d.dataset.ccode || '';
      selfHit = !!k && (codeMode ? code === k : (d.dataset.text || '').includes(k));
      d.querySelectorAll('li').forEach(li => {
        const h = !!k && !codeMode && (li.dataset.text || '').includes(k);
        li.classList.toggle('bhit', h);      /* 命中的商品行标蓝 */
        if (h) subHit = true;
      });
    }
    const show = !k || selfHit || subHit;
    d.style.display = show ? '' : 'none';
    d.open = !!k && show;                  /* 搜索时自动展开命中的卡片 */
  });
  syncBtn();
}

q.addEventListener('input', doSearch);
vP.addEventListener('click', () => { view = 'p'; applyView(); });
vC.addEventListener('click', () => { view = 'c'; applyView(); });

/* 按当前可见项的开合状态刷新按钮文字 */
function syncBtn() {
  const anyClosed = vis().some(d => !d.open);
  tgl.textContent = anyClosed ? '展' : '收';
}

tgl.addEventListener('click', () => {
  const anyClosed = vis().some(d => !d.open);
  vis().forEach(d => { d.open = anyClosed; });
  syncBtn();
});

applyView();     /* 初始化：恢复上次视图并应用当前搜索词 */

/* 说明：标题与总览在吸顶栏之外，随内容自然滚走；吸顶栏用 position:sticky
   常驻吸顶，全程不监听 scroll，所以不会和手机浏览器地址栏伸缩互相触发抖动 */
"""


def _log(quiet, msg):
    if not quiet:
        print(msg)


# ── 统一的工作表读取（.xls 用 xlrd，.xlsx 用 openpyxl，接口一致）──
class _XlsxSheet:
    """把 openpyxl 的工作表包装成 xlrd 风格（nrows/ncols/cell_value）"""

    def __init__(self, path):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        self._grid = [list(r) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
        self.nrows = len(self._grid)
        self.ncols = max((len(r) for r in self._grid), default=0)

    def cell_value(self, r, c):
        if r >= self.nrows or c >= len(self._grid[r]):
            return ""
        v = self._grid[r][c]
        return "" if v is None else v


def open_first_sheet(path):
    """打开第一个工作表：.xlsx 走 openpyxl，其余走 xlrd"""
    if os.path.splitext(path)[1].lower() == ".xlsx":
        return _XlsxSheet(path)
    import xlrd
    return xlrd.open_workbook(path).sheet_by_index(0)


def find_customer(order_no, cust):
    """按与引擎一致的口径找客户：先精确相等，再互相包含
       （引擎的明细匹配就是 contains / 被 contains）"""
    if not order_no:
        return ""
    if order_no in cust:
        return cust[order_no]
    for k, v in cust.items():
        if order_no in k or k in order_no:
            return v
    return ""


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


def sub_rank(big, small):
    """小类在 SUB_ORDER[大类] 中的次序；未配置或未列入的返回 (1, 0) 排后面
       先精确匹配，再取“关键词被小类名包含”中的最长项
       （这样“方盒（大）”这类写法也能归到位）"""
    order = SUB_ORDER.get((big or "").strip())
    if not order:
        return (1, 0)
    key = (small or "").strip()
    for i, k in enumerate(order):
        if key == k:
            return (0, i)
    best = None
    for i, k in enumerate(order):
        if k and k in key:
            cand = (len(k), -i)
            if best is None or cand > best:
                best = cand
    if best is not None:
        return (0, -best[1])
    return (1, 0)


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
    """订单号 → 客户单位（.xls/.xlsx 都支持）"""
    m = {}
    try:
        sh = open_first_sheet(SALES)
    except Exception as e:
        _log(quiet, "!! 打不开销售订单文件：%s" % e)
        return m
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
    cust = load_customers(quiet)

    if not os.path.isdir(DET_DIR):
        return {"error": "找不到目录 %s" % DET_DIR}
    # 同时支持 .xls 与 .xlsx；排序后处理，避免 os.listdir 顺序不定影响分类归属
    files = sorted(f for f in os.listdir(DET_DIR)
                   if f.lower().endswith((".xls", ".xlsx")) and not f.startswith("~$"))
    if not files:
        return {"error": r"D:\明细表格 下没有 .xls/.xlsx 文件"}

    # (商品名称, 规格, 单位) -> {"cat": (大类, 小类), "cust": {客户: 数量}}
    agg = {}
    bad = 0
    miss_cust = 0
    for fn in files:
        # 订单号不再猜"第一段数字"：先在文件名里找出能对上销售订单的号码，
        # 先精确、再退化为"互相包含"（与引擎口径一致），都不中则取最长数字串
        nums = re.findall(r"\d+", fn)
        on = ""
        for n in nums:
            if n in cust:
                on = n
                break
        if not on:
            for n in sorted(nums, key=len, reverse=True):
                if find_customer(n, cust):
                    on = n
                    break
        if not on and nums:
            on = max(nums, key=len)
        cu = find_customer(on, cust)
        if not cu:
            cu = ("未知客户(%s)" % on) if on else "未知客户(无订单号)"
            miss_cust += 1
        try:
            sh = open_first_sheet(os.path.join(DET_DIR, fn))
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
        if bad and bad >= len(files):
            return {"error": "全部 %d 个明细文件都读不出来（格式不支持？缺 xlrd/openpyxl？）" % bad}
        return {"error": "没有汇总到任何商品行（读取失败 %d/%d 个）" % (bad, len(files))}

    # 排序：大类（CAT_ORDER）→ 小类（SUB_ORDER 指定的优先，其余自然序）
    #       → 商品名称 → 规格 → 单位
    items = sorted(agg.items(),
                   key=lambda kv: (cat_rank(kv[1]["cat"][0]),
                                   sub_rank(kv[1]["cat"][0], kv[1]["cat"][1]),
                                   natkey(kv[1]["cat"][1]),
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
        cards.append(
            '<details data-text="%s" style="background:%s"><summary>'
            '<span class="nm">%s</span>'
            '<span class="sub">%s</span>'
            '<span class="unit">%s</span>'
            '<span class="tq">%s</span>'
            '</summary><ul>%s</ul></details>'
            % (esc(text), color_of.get(big, PALETTE[0]), esc(name),
               esc(spec), esc(unit), num(sum(cust.values())), rows)
        )

    # ── 反向视图：客户 → 商品（同一份数据，换个方向看）──
    by_cust = {}
    for (name, spec, unit), v in items:
        for cu, q in v["cust"].items():
            by_cust.setdefault(cu, []).append((name, spec, unit, q, v["cat"][0]))
    cards_c = []
    for cu, rows in sorted(by_cust.items(), key=lambda kv: natkey(kv[0])):
        text = " ".join([cu] + [r[0] for r in rows] + [r[1] for r in rows]).lower()
        code = (re.split(r"[\s★\-_]+", cu)[0] or "").lower()
        lis = "".join(
            '<li data-text="%s">'
            '<span class="nm">%s</span><span class="sub">%s</span>'
            '<span class="unit">%s</span><span class="tq">%s</span></li>'
            % (esc(("%s %s" % (n, s)).lower()), esc(n), esc(s), esc(u), num(q))
            for n, s, u, q, _big in sorted(rows, key=lambda r: (-r[3], natkey(r[0])))
        )
        cards_c.append(
            '<details data-ccode="%s" data-text="%s" style="background:%s"><summary>'
            '<span class="cnm">%s</span>'
            '<span class="tq">%s</span>'
            '</summary><ul>%s</ul></details>'
            % (esc(code), esc(text), "#eefaf1",     # 客户栏统一淡绿
               esc(cu), num(sum(r[3] for r in rows)), lis))

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
    <button class="vbtn on" id="vP">商</button>
    <button class="vbtn" id="vC">客</button>
    <input id="q" placeholder="搜索商品 / 客户…">
    <button id="tgl">展</button>
  </div>
</header>
<div id="list">%s</div>
<div id="listC" hidden>%s</div>
<footer>点标题行可展开明细；「商」看每个商品卖给了谁，「客」看每家店拿了什么</footer>
<script>%s</script></body></html>""" % (
        now, CSS, len(items), data["cust_n"], num(data["total"]),
        data["files"], now, warn, "".join(cards), "".join(cards_c), JS)


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
