#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v3.1.3 引擎回归验证脚本

用法（在 D:\\WFR\\D2Y-pyqt 目录下）：
    python3 verify_engine.py

它会加载同目录的 jinhua_engine.dll，用真实数据跑 4 个场景并打印诊断：

  1) 正常路径                      force=0 → 期望 matched == orders 且 fatal=false
  2) 缺 销售订单.xls               force=0 → 期望 fatal=true 且「不生成」新文件
  3) 明细为空（源明细与 MXBG2 都移走）force=0 → 期望「匹配明细 0 单」致命中止、不生成文件
  4) 同场景 force=1                → 期望 fatal=false、生成文件、并记录对应 warning

脚本只移动 D:\\明细表格 下的源文件与 MXBG2 子目录，最后会恢复（try/finally），
并且校验恢复后的文件数量；delete 参数固定为 0，不会删除任何源文件。
"""
import ctypes
import json
import os
import shutil
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

DLL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "jinhua_engine.dll")
OUT = r"D:\订单表格"
DET = r"D:\明细表格"
SALES = os.path.join(OUT, "销售订单.xls")
MX = os.path.join(DET, "MXBG2")
HOLD_D = r"D:\明细表格__hold_test"
HOLD_MX = r"D:\明细表格__MXBG2_test"

if not os.path.exists(DLL):
    print("!! 未找到引擎 DLL:", DLL)
    sys.exit(1)

lib = ctypes.CDLL(DLL)
CB = ctypes.CFUNCTYPE(None, ctypes.c_uint32, ctypes.c_char_p)
lib.engine_process.argtypes = [ctypes.c_uint8, ctypes.c_uint8, ctypes.c_uint8, CB]
lib.engine_process.restype = ctypes.c_void_p
lib.engine_free.argtypes = [ctypes.c_void_p]


def run(force):
    @CB
    def cb(pct, lp):  # 不关心进度
        pass
    ptr = lib.engine_process(0, 0, force, cb)   # delete=0, stats=0
    raw = ctypes.cast(ptr, ctypes.c_char_p).value.decode("utf-8")
    lib.engine_free(ptr)
    return json.loads(raw)


def show(tag, r):
    print("\n=== %s ===" % tag)
    print("  error=%r  fatal=%s  file=%r" % (r.get("error"), r.get("fatal"), r.get("file")))
    print("  orders=%s matched=%s customers=%s products=%s"
          % (r.get("orders"), r.get("matched"), r.get("customers"), r.get("products")))
    for d in r.get("errors", []):
        print("  [ERROR] %s | %s | %s" % (d["field"], d["problem"], d["detail"][:230]))
    for d in r.get("warnings", []):
        print("  [warn ] %s | %s | %s" % (d["field"], d["problem"], d["detail"][:190]))


def outs():
    if not os.path.exists(OUT):
        return set()
    return set(f for f in os.listdir(OUT)
               if f.startswith("优路达导入模板") and f.endswith(".xlsx"))


for p in (HOLD_D, HOLD_MX):
    if os.path.exists(p):
        shutil.rmtree(p)

print("[dll]", DLL)

# 1) 正常路径
b = outs()
r = run(0)
show("1. 正常路径 force=0", r)
print("  新生成:", sorted(outs() - b))
n_orders, n_matched = r.get("orders"), r.get("matched")
print("  结论:", "PASS" if (r.get("fatal") is False and n_orders and n_orders == n_matched) else "需要人工确认")

# 2) 缺销售订单
b = outs()
shutil.move(SALES, SALES + ".holdtest")
try:
    r = run(0)
    show("2. 缺 销售订单.xls force=0（应 fatal 且不生成文件）", r)
    print("  新生成:", sorted(outs() - b), "(期望空)")
finally:
    shutil.move(SALES + ".holdtest", SALES)
print("  销售订单已恢复:", os.path.exists(SALES), os.path.getsize(SALES), "字节")

# 3) 明细为空：force=0 应中止；force=1 应继续
os.makedirs(HOLD_D, exist_ok=True)
srcs = [f for f in os.listdir(DET)
        if f.lower().endswith((".xls", ".xlsx")) and not f.startswith("~$")]
for f in srcs:
    shutil.move(os.path.join(DET, f), os.path.join(HOLD_D, f))
if os.path.exists(MX):
    shutil.move(MX, HOLD_MX)
try:
    b = outs()
    r = run(0)
    show("3. 明细为空 force=0（应中止、不生成文件）", r)
    print("  新生成:", sorted(outs() - b), "(期望空)")

    b = outs()
    r = run(1)
    show("4. 明细为空 force=1（应继续生成并记 warning）", r)
    print("  新生成:", sorted(outs() - b))
finally:
    if os.path.exists(MX):
        shutil.rmtree(MX)
    if os.path.exists(HOLD_MX):
        shutil.move(HOLD_MX, MX)
    for f in os.listdir(HOLD_D):
        shutil.move(os.path.join(HOLD_D, f), os.path.join(DET, f))
    shutil.rmtree(HOLD_D, ignore_errors=True)

restored = len([f for f in os.listdir(DET)
                if f.lower().endswith((".xls", ".xlsx")) and not f.startswith("~$")])
print("\n恢复校验: 源明细 %d 个 | MXBG2 存在=%s 文件数=%s"
      % (restored, os.path.exists(MX), len(os.listdir(MX)) if os.path.exists(MX) else -1))
