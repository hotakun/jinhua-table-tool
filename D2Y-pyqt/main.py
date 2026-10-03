#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金华聚火表格处理 - PySide6 版
PySide6 负责 GUI，Rust DLL 负责所有 Excel 处理
与 ttkbootstrap 版的对比练习项目
"""
import sys
import os
import json
import ctypes
import threading
import time
import sqlite3
import ssl
import urllib.request
import urllib.error
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QCheckBox, QGroupBox, QProgressBar,
    QTextEdit, QTableWidget, QTableWidgetItem, QTabWidget,
    QSplashScreen, QFrame, QToolButton, QStatusBar, QMessageBox,
    QHeaderView, QGraphicsDropShadowEffect,
)
from PySide6.QtCore import (
    Qt, QPoint, QUrl, QThread, Signal, QTimer, QPropertyAnimation, QEasingCurve,
)
from PySide6.QtGui import (
    QFont, QIcon, QPixmap, QMovie, QColor, QPalette,
    QShortcut, QKeySequence, QDesktopServices,
)

# ============================================================================
# 常量
# ============================================================================
CURRENT_VERSION = "3.2.1"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 操作日志数据库
LOG_DB_DIR = r"D:\订单表格\logs"
LOG_DB_PATH = os.path.join(LOG_DB_DIR, "operations.db")
# 更新检查 / 下载（v3.1.3 新增）
UPDATE_API = "https://api.github.com/repos/hotakun/jinhua-table-tool/releases/latest"
UPDATE_PAGE = "https://github.com/hotakun/jinhua-table-tool/releases/latest"
UPDATE_DIR = r"D:\订单表格\Update"

# ============================================================================
# 工具函数（与原版复用逻辑）
# ============================================================================
def find_engine():
    paths = [
        os.path.join(BASE_DIR, "jinhua_engine.dll"),
        os.path.join(BASE_DIR, "_internal", "jinhua_engine.dll"),
        os.path.join(BASE_DIR, "..", "Rust", "target", "release", "jinhua_engine.dll"),
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None

# ── 操作日志数据库 ──
def init_db():
    """初始化数据库和表（首次调用时自动创建）"""
    os.makedirs(LOG_DB_DIR, exist_ok=True)
    conn = sqlite3.connect(LOG_DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS operations (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            op_time    TEXT    NOT NULL,
            operator   TEXT    NOT NULL,
            status     TEXT    NOT NULL,
            orders     INTEGER,
            customers  INTEGER,
            products   INTEGER,
            matched    INTEGER,
            output_file TEXT,
            stats_enabled INTEGER,
            delete_enabled INTEGER,
            error_msg  TEXT,
            duration_ms INTEGER
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS diagnostics (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            op_id      INTEGER,
            op_time    TEXT,
            level      TEXT,
            field      TEXT,
            problem    TEXT,
            detail     TEXT
        )
    """)
    conn.commit()
    conn.close()

def log_operation(op_time, operator, status, orders=0, customers=0,
                  products=0, matched=0, output_file="",
                  stats_enabled=0, delete_enabled=0,
                  error_msg="", duration_ms=0):
    """写入一条操作日志，返回该记录 id（失败返回 0，不影响主流程）"""
    try:
        os.makedirs(LOG_DB_DIR, exist_ok=True)
        conn = sqlite3.connect(LOG_DB_PATH)
        cur = conn.execute("""
            INSERT INTO operations
                (op_time, operator, status, orders, customers,
                 products, matched, output_file, stats_enabled,
                 delete_enabled, error_msg, duration_ms)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (op_time, operator, status, orders, customers,
              products, matched, output_file, stats_enabled,
              delete_enabled, error_msg, duration_ms))
        conn.commit()
        rid = cur.lastrowid or 0
        conn.close()
        return rid
    except Exception:
        return 0

def log_diagnostics(op_id, errors=(), warnings=(), op_time=""):
    """把诊断明细写入 diagnostics 表，便于事后按字段检索（失败静默）"""
    if not op_id:
        return
    try:
        os.makedirs(LOG_DB_DIR, exist_ok=True)
        conn = sqlite3.connect(LOG_DB_PATH)
        rows = [("error", d) for d in errors] + [("warning", d) for d in warnings]
        for level, d in rows:
            conn.execute(
                "INSERT INTO diagnostics (op_id, op_time, level, field, problem, detail)"
                " VALUES (?, ?, ?, ?, ?, ?)",
                (op_id, op_time, level, d.get("field", ""),
                 d.get("problem", ""), d.get("detail", "")))
        conn.commit()
        conn.close()
    except Exception:
        pass


# ── TXT 对齐工具（与 Rust 引擎 _dwidth/_pad 完全一致: 中文=4, ASCII=2）──
def _dwidth(s):
    w = 0
    for c in s:
        o = ord(c)
        if o in (0x200B, 0x200C, 0x200D, 0xFEFF):
            continue
        w += 4 if o > 0x2E80 else 2
    return w

def _numstr(v):
    """数量显示：整数不带小数，否则保留两位"""
    try:
        f = float(v)
    except Exception:
        return str(v)
    return str(int(f)) if f.is_integer() else ("%.2f" % f)


def _pad(s, target_w):
    """按显示宽度补齐（中文=4、ASCII=2）。
       关键：1 个空格只占 2 个宽度单位，所以补空格数 = 缺口 / 2；
       超宽则按宽度截断并补 "…"，避免把后面的列顶歪。"""
    w = _dwidth(s)
    if w <= target_w:
        return s + ' ' * max(0, (target_w - w) // 2)
    limit = max(0, target_w - 2)
    out, used = "", 0
    for ch in s:
        cw = 4 if ord(ch) > 0x2E80 else 2
        if used + cw > limit:
            break
        out += ch
        used += cw
    out += "…"
    used += 2
    if used < target_w:
        out += ' ' * ((target_w - used) // 2)
    return out

# Rust 引擎统计列宽度: name=46, spec=36, unit=10, qty(右对齐)=6
STATS_W = (46, 36, 10, 6)


# ── 版本比较与更新检查（v3.1.3 新增）──
def _ver_key(v):
    """把 'v3.1.2' / 'V2.7.0' 这类版本号转成可比较的数值元组（大小写 v 都容错）"""
    parts = []
    for seg in (v or "").strip().lstrip("vV").split("."):
        digits = "".join(ch for ch in seg if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)

def _ver_newer(a, b):
    return _ver_key(a) > _ver_key(b)

_ssl_ctx = None

def _fetch(url, timeout=15):
    """带 SSL 容错的 GET"""
    global _ssl_ctx
    req = urllib.request.Request(url, headers={"User-Agent": "JinhuaJuhuo"})
    try:
        return urllib.request.urlopen(req, timeout=timeout)
    except ssl.SSLError:
        if _ssl_ctx is None:
            _ssl_ctx = ssl.create_default_context()
            _ssl_ctx.check_hostname = False
            _ssl_ctx.verify_mode = ssl.CERT_NONE
        return urllib.request.urlopen(req, timeout=timeout, context=_ssl_ctx)

def check_update():
    """检查更新，返回 (status, info)：
         ("new",    {版本信息})  有新版本
         ("latest", None)        已是最新
         ("fail",   None)        网络或接口异常
    """
    try:
        with _fetch(UPDATE_API) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        tag = (data.get("tag_name") or "").strip()
        if not tag or not _ver_newer(tag, CURRENT_VERSION):
            return ("latest", None)
        for a in data.get("assets", []):
            name = a.get("name", "") or ""
            if name.lower().endswith(".exe") and "Setup" in name:
                return ("new", {
                    "version": tag.lstrip("vV"),
                    "name": name,
                    "url": a.get("browser_download_url") or "",
                    "size": int(a.get("size") or 0),
                    "page": UPDATE_PAGE,
                })
        return ("latest", None)
    except Exception:
        return ("fail", None)

def _diags_text(data):
    """把引擎返回的 errors/warnings 拼成可读文本（含出错字段、原因与定位信息）"""
    out = []
    for tag, label in (("errors", "错误"), ("warnings", "警告")):
        items = (data or {}).get(tag) or []
        if not items:
            continue
        out.append("【%s】" % label)
        for i, d in enumerate(items, 1):
            out.append("%d. %s — %s" % (i, d.get("field", ""), d.get("problem", "")))
            det = (d.get("detail") or "").strip()
            if det:
                out.append("    " + det)
        out.append("")
    return "\n".join(out).strip() or "（无诊断信息）"

# ============================================================================
# QSS 样式表
# ============================================================================
LIGHT_QSS = """
/* 全局 */
QMainWindow, QWidget {
    font-family: "Microsoft YaHei";
    font-size: 11pt;
}

/* 按钮 */
QPushButton {
    border-radius: 8px;
    padding: 10px 28px;
    font-size: 14pt;
    font-weight: bold;
    min-height: 20px;
}
QPushButton#executeBtn {
    background-color: #28a745;
    color: white;
    border: none;
}
QPushButton#executeBtn:hover {
    background-color: #218838;
}
QPushButton#executeBtn:pressed {
    background-color: #1e7e34;
}
QPushButton#executeBtn:disabled {
    background-color: #6c757d;
    color: #ccc;
}
QPushButton#closeBtn {
    background-color: #dc3545;
    color: white;
    border: none;
}
QPushButton#closeBtn:hover {
    background-color: #c82333;
}
QPushButton#openBtn {
    background-color: #28a745;
    color: white;
    border: none;
}
QPushButton#openBtn:hover {
    background-color: #218838;
}

/* 进度条 */
QProgressBar {
    border: 1px solid #ddd;
    border-radius: 6px;
    text-align: center;
    background-color: #f0f0f0;
    height: 22px;
}
QProgressBar::chunk {
    background-color: #28a745;
    border-radius: 5px;
}

/* GroupBox 卡片 */
QGroupBox {
    border: 1px solid #e0e0e0;
    border-radius: 10px;
    margin-top: 12px;
    padding-top: 16px;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 8px;
}

/* 表格 */
QTableWidget {
    background-color: #ffffff;
    color: #1e1e1e;
    gridline-color: #e0e0e0;
    alternate-background-color: #f8f9fa;
    selection-background-color: #cce5ff;
}
QTableWidget::item {
    color: #1e1e1e;
}
QHeaderView::section {
    background-color: #f0f0f0;
    color: #1e1e1e;
    padding: 6px;
    border: none;
    border-bottom: 2px solid #ddd;
    font-weight: bold;
}

/* Tab */
QTabWidget::pane {
    border: 1px solid #e0e0e0;
    border-radius: 6px;
    background-color: #ffffff;
}
QTabBar::tab {
    background: #f0f0f0;
    color: #1e1e1e;
    padding: 8px 16px;
    border: 1px solid #e0e0e0;
    border-bottom: none;
    border-radius: 6px 6px 0 0;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #1e1e1e;
}

/* ToolButton 暗色模式切换 */
QToolButton#themeBtn {
    border: none;
    border-radius: 12px;
    padding: 4px;
    font-size: 16pt;
}
QToolButton#themeBtn:hover {
    background-color: #e0e0e0;
}

QTextEdit {
    background-color: #ffffff;
    color: #1e1e1e;
    border: 1px solid #ddd;
}

QCheckBox {
    color: #1e1e1e;
}

QLabel {
    color: #1e1e1e;
}

QStatusBar {
    color: #999;
    background-color: #f0f0f0;
}

QLabel#subtitleLabel {
    color: #888;
}

QScrollBar:vertical {
    background: #f0f0f0;
    width: 10px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #c0c0c0;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""

DARK_QSS = """
QMainWindow, QWidget {
    font-family: "Microsoft YaHei";
    font-size: 11pt;
    color: #e0e0e0;
    background-color: #2b2b2b;
}

QPushButton {
    border-radius: 8px;
    padding: 10px 28px;
    font-size: 14pt;
    font-weight: bold;
    min-height: 20px;
}
QPushButton#executeBtn {
    background-color: #28a745;
    color: white;
    border: none;
}
QPushButton#executeBtn:hover {
    background-color: #34ce57;
}
QPushButton#executeBtn:disabled {
    background-color: #555;
    color: #999;
}
QPushButton#closeBtn {
    background-color: #dc3545;
    color: white;
    border: none;
}
QPushButton#closeBtn:hover {
    background-color: #e4606d;
}
QPushButton#openBtn {
    background-color: #28a745;
    color: white;
    border: none;
}
QPushButton#openBtn:hover {
    background-color: #34ce57;
}

QProgressBar {
    border: 1px solid #555;
    border-radius: 6px;
    text-align: center;
    background-color: #3a3a3a;
    color: #e0e0e0;
    height: 22px;
}
QProgressBar::chunk {
    background-color: #28a745;
    border-radius: 5px;
}

QGroupBox {
    border: 1px solid #555;
    border-radius: 10px;
    margin-top: 12px;
    padding-top: 16px;
    font-weight: bold;
    color: #e0e0e0;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 8px;
    color: #e0e0e0;
}

QTableWidget {
    gridline-color: #444;
    alternate-background-color: #333;
    selection-background-color: #1a5276;
    color: #e0e0e0;
    background-color: #2b2b2b;
}
QTableWidget::item {
    color: #e0e0e0;
}
QHeaderView::section {
    background-color: #3a3a3a;
    padding: 6px;
    border: none;
    border-bottom: 2px solid #555;
    font-weight: bold;
    color: #e0e0e0;
}

QTabWidget::pane {
    border: 1px solid #555;
    border-radius: 6px;
    background-color: #2b2b2b;
}
QTabBar::tab {
    background: #3a3a3a;
    color: #999;
    padding: 8px 16px;
    border: 1px solid #555;
    border-bottom: none;
    border-radius: 6px 6px 0 0;
}
QTabBar::tab:selected {
    background: #2b2b2b;
    color: #e0e0e0;
}

QTextEdit {
    background-color: #2b2b2b;
    color: #e0e0e0;
    border: 1px solid #555;
}

QCheckBox {
    color: #e0e0e0;
}

QLabel {
    color: #e0e0e0;
}

QToolButton#themeBtn {
    border: none;
    border-radius: 12px;
    padding: 4px;
    font-size: 16pt;
    color: #e0e0e0;
}
QToolButton#themeBtn:hover {
    background-color: #444;
}

QStatusBar {
    color: #999;
    background-color: #222;
}

QLabel#subtitleLabel {
    color: #999;
}

QScrollBar:vertical {
    background: #2b2b2b;
    width: 10px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #555;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""

# ============================================================================
# 工作线程 — Rust DLL 调用
# ============================================================================
class EngineWorker(QThread):
    progress_updated = Signal(int, str, str)   # pct, label, percentage_str
    stats_received = Signal(int, list)          # total, lines
    engine_finished = Signal(dict)              # result_json
    engine_error = Signal(str, dict)            # error message + 完整结果(含诊断)

    def __init__(self, dll_path, delete_files, show_stats, force=False, parent=None):
        super().__init__(parent)
        self.dll_path = dll_path
        self.delete_files = delete_files
        self.show_stats = show_stats
        self.force = force          # True = 忽略致命列识别问题继续生成（「仍然生成」）

    def run(self):
        try:
            for d in [r'D:\订单表格', r'D:\明细表格', r'D:\明细表格\MXBG2', UPDATE_DIR]:
                os.makedirs(d, exist_ok=True)

            # 先把商品汇总需要的数据读出来（引擎随后可能按勾选清空 D:\明细表格，
            # 清完就再也算不出「商品 × 客户」了）
            self.summary_data = None
            try:
                import summary
                self.summary_data = summary.collect(quiet=True)
            except Exception:
                self.summary_data = None

            lib = ctypes.CDLL(self.dll_path)
            CB = ctypes.CFUNCTYPE(None, ctypes.c_uint32, ctypes.c_char_p)

            @CB
            def progress_cb(pct, label_ptr):
                try:
                    label = label_ptr.decode('utf-8') if label_ptr else ""
                    if label.startswith("STATS:"):
                        parts = label[6:].split("|", 1)
                        total = int(parts[0])
                        lines = json.loads(f"[{parts[1]}]") if len(parts) > 1 else []
                        self.stats_received.emit(total, lines)
                    else:
                        self.progress_updated.emit(pct, label, f"{pct}%")
                except Exception:
                    pass  # 单条 stats 数据解析失败不影响整体

            lib.engine_process.argtypes = [ctypes.c_uint8, ctypes.c_uint8, ctypes.c_uint8, CB]
            lib.engine_process.restype = ctypes.c_void_p
            lib.engine_free.argtypes = [ctypes.c_void_p]

            self.progress_updated.emit(5, "引擎处理中...", "5%")
            delete = 1 if self.delete_files else 0
            stats = 1 if self.show_stats else 0
            force = 1 if self.force else 0
            ptr = lib.engine_process(delete, stats, force, progress_cb)
            result_json = ctypes.cast(ptr, ctypes.c_char_p).value.decode('utf-8')
            lib.engine_free(ptr)

            result_data = json.loads(result_json)
            result_data["_summary"] = self.summary_data   # 供「打开并退出」时生成 HTML
            if "error" in result_data:
                self.engine_error.emit(result_data["error"], result_data)
            else:
                self.engine_finished.emit(result_data)

        except OSError as e:
            self.engine_error.emit(f"Rust引擎DLL加载失败: {str(e)}", {})
        except Exception as e:
            self.engine_error.emit(f"引擎异常: {str(e)}", {})


# ============================================================================
# 主窗口
# ============================================================================
class JinhuaPyQtApp(QMainWindow):
    # 后台线程 → 主线程的结果回传（QTimer 在没有事件循环的子线程里不会触发，
    # 必须用 Qt 信号跨线程排队，否则界面收不到任何更新结果）
    update_checked = Signal(str, object)             # status(new/latest/fail), info
    update_progress = Signal(int)                   # 下载百分比
    update_downloaded = Signal(str, bool, bool)     # dest, size_ok, pe_ok
    update_failed = Signal(str)                     # 错误信息

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"金华聚火表格处理 v{CURRENT_VERSION}")
        self.setGeometry(100, 100, 660, 820)
        self.setMinimumSize(580, 680)

        # 图标
        for p in [os.path.join(BASE_DIR, "favicon.ico"),
                  os.path.join(BASE_DIR, "_internal", "favicon.ico")]:
            if os.path.exists(p):
                self.setWindowIcon(QIcon(p))
                break

        self.is_running = False

        # 初始化操作日志数据库
        init_db()
        self.is_dark = False
        self.result_data = None
        self.detail_stats_lines = []
        self.detail_stats_total = 0
        self.output_file_path = ""
        self.update_info = None       # 发现的新版本信息
        self.is_downloading = False
        self._manual_check = False    # 区分启动静默检查与手动点击检查

        self._setup_ui()
        self._setup_statusbar()
        self._setup_shortcuts()

        # 检查/下载期间按钮转圈，让用户看得出程序在动
        self._spin_frames = ["◐", "◓", "◑", "◒"]
        self._spin_idx = 0
        self._spin_timer = QTimer(self)
        self._spin_timer.setInterval(180)
        self._spin_timer.timeout.connect(self._spin_step)

        # 跨线程信号接到主线程槽
        self.update_checked.connect(self._on_update_checked)
        self.update_progress.connect(self._on_update_progress)
        self.update_downloaded.connect(self._on_update_downloaded)
        self.update_failed.connect(self._on_update_failed)

        # 启动 1.5s 后静默检查更新：发现新版只在按钮/状态栏提示，无网或失败不提示
        QTimer.singleShot(1500, self._check_update_silent)

    # ── UI 构建 ──────────────────────────────────────────────
    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        self.main_layout = QVBoxLayout(central)
        self.main_layout.setContentsMargins(28, 18, 28, 18)
        self.main_layout.setSpacing(10)

        # ── 顶部标题 ──
        self._build_header()


        # ── 复选框 ──
        self._build_checkboxes()

        # ── 执行按钮 ──
        self._build_execute_btn()

        # ── 进度区 ──
        self._build_progress_group()

        # ── 结果区 (Tab) ──
        self._build_result_tabs()

        # ── 完成按钮行（初始隐藏） ──
        self._build_done_buttons()

        self.main_layout.addStretch()
        self._apply_qss()
        self._apply_light_palette()

    def _apply_light_palette(self):
        """启动时强制亮色调色板，不受系统暗色主题影响"""
        QApplication.setStyle("Fusion")
        p = QPalette()
        p.setColor(QPalette.Window, QColor(240, 240, 240))
        p.setColor(QPalette.WindowText, QColor(30, 30, 30))
        p.setColor(QPalette.Base, QColor(255, 255, 255))
        p.setColor(QPalette.AlternateBase, QColor(245, 245, 245))
        p.setColor(QPalette.Text, QColor(30, 30, 30))
        p.setColor(QPalette.Button, QColor(240, 240, 240))
        p.setColor(QPalette.ButtonText, QColor(30, 30, 30))
        p.setColor(QPalette.Highlight, QColor(40, 167, 69))
        p.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        QApplication.setPalette(p)

    def _build_header(self):
        h_layout = QHBoxLayout()
        h_layout.addStretch()

        # 更新按钮（在 ? 左边）
        self.update_btn = QToolButton()
        self.update_btn.setObjectName("themeBtn")
        self.update_btn.setText("↻")
        self.update_btn.setToolTip("检查更新")
        self.update_btn.setFixedSize(32, 32)
        self.update_btn.clicked.connect(self._on_update_clicked)

        # 帮助按钮（右上角，? 旁边是暗色模式按钮）
        self.help_btn = QToolButton()
        self.help_btn.setObjectName("themeBtn")
        self.help_btn.setText("?")
        self.help_btn.setToolTip("查看文件路径说明")
        self.help_btn.setFixedSize(32, 32)
        self.help_btn.clicked.connect(self._show_path_popup)

        self.theme_btn = QToolButton()
        self.theme_btn.setObjectName("themeBtn")
        self.theme_btn.setText("☽")
        self.theme_btn.setToolTip("切换亮色模式" if self.is_dark else "切换暗色模式")
        self.theme_btn.setFixedSize(32, 32)
        self.theme_btn.clicked.connect(self._toggle_theme)

        title_layout = QVBoxLayout()
        self.title_label = QLabel(f"金华聚火表格处理 v{CURRENT_VERSION}")
        self.title_label.setFont(QFont("Microsoft YaHei", 20, QFont.Bold))

        self.subtitle_label = QLabel("订小易_转_优路达_表格处理")
        self.subtitle_label.setFont(QFont("Microsoft YaHei", 10))
        self.subtitle_label.setObjectName("subtitleLabel")

        title_layout.addWidget(self.title_label, alignment=Qt.AlignCenter)
        title_layout.addWidget(self.subtitle_label, alignment=Qt.AlignCenter)

        h_layout.addLayout(title_layout)
        h_layout.addStretch()
        h_layout.addWidget(self.update_btn, alignment=Qt.AlignTop)
        h_layout.addWidget(self.help_btn, alignment=Qt.AlignTop)
        h_layout.addWidget(self.theme_btn, alignment=Qt.AlignTop)
        self.main_layout.addLayout(h_layout)

    # ── 更新检查与下载（v3.1.3）──────────────────────────────
    def _start_spin(self):
        """按钮转圈，表示正在检查/下载"""
        self._spin_idx = 0
        self.update_btn.setText(self._spin_frames[0])
        self._spin_timer.start()

    def _spin_step(self):
        self._spin_idx = (self._spin_idx + 1) % len(self._spin_frames)
        self.update_btn.setText(self._spin_frames[self._spin_idx])

    def _stop_spin(self):
        self._spin_timer.stop()
        self.update_btn.setText("↓" if self.update_info else "↻")

    def _check_update_silent(self):
        """启动时静默检查：只在发现新版时给按钮加提示，失败/无网一律不打扰"""
        self._manual_check = False
        threading.Thread(target=lambda: self.update_checked.emit(*check_update()),
                         daemon=True).start()

    def _on_update_clicked(self):
        """点击 ↻ / ↓：手动检查更新（必须有明确反馈）"""
        if self.is_downloading:
            return
        self._manual_check = True
        self.status_label_sb.setText("正在检查更新...")
        self.update_btn.setEnabled(False)
        self._start_spin()
        threading.Thread(target=lambda: self.update_checked.emit(*check_update()),
                         daemon=True).start()

    def _on_update_checked(self, status, info):
        """主线程槽：后台检查完成后回到这里（静默检查只在有新版时提示）"""
        if not self._manual_check:
            if status == "new" and info:
                self._mark_update_available(info)
            return
        self._after_check(status, info)

    def _mark_update_available(self, info):
        self.update_info = info
        self.update_btn.setText("↓")
        self.update_btn.setToolTip("有新版本 v%s，点击更新" % info["version"])
        self.status_label_sb.setText("有新版本 v%s（点 ↓ 更新）" % info["version"])

    def _after_check(self, status, info):
        self.update_btn.setEnabled(True)
        self._stop_spin()
        # 网络/接口异常：与“已是最新”区分开
        if status == "fail":
            self.status_label_sb.setText("检查更新失败")
            box = QMessageBox(self)
            box.setWindowTitle("检查更新")
            box.setIcon(QMessageBox.Warning)
            box.setText("检查更新失败，请检查网络后重试。")
            page_btn = box.addButton("打开下载页面", QMessageBox.ActionRole)
            box.addButton("关闭", QMessageBox.RejectRole)
            box.exec()
            if box.clickedButton() is page_btn:
                QDesktopServices.openUrl(QUrl(UPDATE_PAGE))
            return
        if status != "new" or not info:
            if self.update_info:
                info = self.update_info
            else:
                self.status_label_sb.setText("已是最新版本")
                QMessageBox.information(self, "检查更新",
                    "当前已是最新版本 v%s" % CURRENT_VERSION)
                return
        self.update_info = info
        self._mark_update_available(info)
        box = QMessageBox(self)
        box.setWindowTitle("发现新版本")
        box.setIcon(QMessageBox.Information)
        box.setText("发现新版本 v%s，是否下载安装？" % info["version"])
        dl_btn = box.addButton("下载并安装", QMessageBox.AcceptRole)
        box.addButton("取消", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is dl_btn:
            self._download_update(info)

    def _download_update(self, info):
        if not info.get("url"):
            QDesktopServices.openUrl(QUrl(info["page"]))
            return
        self.is_downloading = True
        self.update_btn.setEnabled(False)
        self._start_spin()
        os.makedirs(UPDATE_DIR, exist_ok=True)
        # 清掉以前下载的安装包（每个 ~37MB，不清会越积越多）
        try:
            for fn in os.listdir(UPDATE_DIR):
                if fn.lower().endswith(".exe"):
                    os.remove(os.path.join(UPDATE_DIR, fn))
        except Exception:
            pass
        dest = os.path.join(UPDATE_DIR, info["name"])

        def work():
            last = ""
            for attempt in range(3):          # 网络抖动时最多重试 3 次
                try:
                    with _fetch(info["url"], timeout=60) as resp, open(dest, "wb") as f:
                        total = int(resp.headers.get("content-length") or 0)
                        got = 0
                        while True:
                            chunk = resp.read(65536)
                            if not chunk:
                                break
                            f.write(chunk)
                            got += len(chunk)
                            if total > 0:
                                self.update_progress.emit(int(got * 100 / total))
                    size_ok = (not info.get("size")) or os.path.getsize(dest) == info["size"]
                    with open(dest, "rb") as f:
                        pe_ok = f.read(2) == b"MZ"
                    self.update_downloaded.emit(dest, size_ok, pe_ok)
                    return
                except Exception as e:
                    last = str(e)
                    time.sleep(1.5 * (attempt + 1))
            self.update_failed.emit(last)

        threading.Thread(target=work, daemon=True).start()

    # ── 下载过程中的三个主线程槽 ──
    def _on_update_progress(self, pct):
        self.status_label_sb.setText("正在下载更新... %d%%" % pct)

    def _on_update_downloaded(self, dest, size_ok, pe_ok):
        self._after_download(dest, size_ok, pe_ok)

    def _on_update_failed(self, err):
        self._download_failed(err)

    def _after_download(self, dest, size_ok, pe_ok):
        self.is_downloading = False
        self.update_btn.setEnabled(True)
        self._stop_spin()
        if not size_ok or not pe_ok:
            try:
                os.remove(dest)
            except Exception:
                pass
            self.status_label_sb.setText("下载失败，请重试")
            QMessageBox.warning(self, "更新失败", "下载的文件不完整，请重试。")
            return
        self.status_label_sb.setText("更新包已下载")
        box = QMessageBox(self)
        box.setWindowTitle("下载完成")
        box.setIcon(QMessageBox.Information)
        box.setText("已下载完成，现在安装吗？")
        box.setInformativeText("安装前本程序会自动退出。")
        now_btn = box.addButton("立即安装", QMessageBox.AcceptRole)
        box.addButton("稍后", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is now_btn:
            try:
                os.startfile(dest)
            except Exception:
                QDesktopServices.openUrl(QUrl.fromLocalFile(dest))
            self.close()

    def _download_failed(self, err, info=None):
        self.is_downloading = False
        self.update_btn.setEnabled(True)
        self._stop_spin()
        self.status_label_sb.setText("下载失败，请重试")
        page = (info or {}).get("page") or (self.update_info or {}).get("page") or UPDATE_PAGE
        box = QMessageBox(self)
        box.setWindowTitle("下载失败")
        box.setIcon(QMessageBox.Warning)
        box.setText("更新包下载失败，请稍后重试。")
        page_btn = box.addButton("打开下载页面", QMessageBox.ActionRole)
        box.addButton("关闭", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is page_btn:
            QDesktopServices.openUrl(QUrl(page))

    def _show_path_popup(self):
        """在 ? 按钮下方弹出路径说明浮层，点击外部自动关闭。"""
        popup = QFrame(self, Qt.Popup | Qt.FramelessWindowHint)
        popup.setObjectName("pathPopup")
        popup.setStyleSheet("""
            QFrame#pathPopup {
                background-color: #ffffff;
                border: 1px solid #ddd;
                border-radius: 10px;
            }
        """ if not self.is_dark else """
            QFrame#pathPopup {
                background-color: #3a3a3a;
                border: 1px solid #555;
                border-radius: 10px;
            }
        """)
        layout = QVBoxLayout(popup)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(4)
        header = QLabel("必须拥有以下文件及目录：")
        header.setFont(QFont("Microsoft YaHei", 9, QFont.Bold))
        header.setStyleSheet("color: #333;" if not self.is_dark else "color: #ddd;")
        layout.addWidget(header)
        for i, path in enumerate([
            "1、D:\\订单表格\\销售订单.xls",
            "2、D:\\订单表格\\客户经纬度.xls",
            "3、D:\\订单表格\\优路达导入模板.xlsx",
            "4、D:\\明细表格\\销售订单详细 - *.xls",
        ], 1):
            lb = QLabel(path)
            lb.setFont(QFont("Microsoft YaHei", 9))
            lb.setStyleSheet("color: #555;" if not self.is_dark else "color: #bbb;")
            layout.addWidget(lb)
        popup.adjustSize()
        # 定位在 ? 按钮下方，右对齐
        pos = self.help_btn.mapToGlobal(QPoint(0, self.help_btn.height() + 4))
        popup.move(pos.x() + self.help_btn.width() - popup.width(), pos.y())
        popup.show()


    def _build_checkboxes(self):
        layout = QVBoxLayout()
        layout.setSpacing(4)
        self.delete_cb = QCheckBox("执行完毕后删除 D:\\明细表格 目录")
        self.delete_cb.setChecked(True)
        self.stats_cb = QCheckBox("查看商品明细数量")
        self.stats_cb.setChecked(True)
        layout.addWidget(self.delete_cb)
        layout.addWidget(self.stats_cb)
        self.main_layout.addLayout(layout)

    def _build_execute_btn(self):
        self.btn_container = QVBoxLayout()
        self.execute_btn = QPushButton("执行表格转换")
        self.execute_btn.setObjectName("executeBtn")
        self.execute_btn.setFixedHeight(48)
        self.execute_btn.clicked.connect(self._execute)
        self.btn_container.addWidget(self.execute_btn, alignment=Qt.AlignCenter)
        self.main_layout.addLayout(self.btn_container)

    def _build_progress_group(self):
        group = QGroupBox("进度")
        self._add_shadow(group)
        layout = QVBoxLayout(group)
        layout.setContentsMargins(14, 18, 14, 12)

        self.progress_label = QLabel("就绪")
        self.progress_label.setFont(QFont("Microsoft YaHei", 11))
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)

        self.pct_label = QLabel("0%")
        self.pct_label.setFont(QFont("Microsoft YaHei", 10))
        layout.addWidget(self.pct_label, alignment=Qt.AlignCenter)

        self.main_layout.addWidget(group)

    def _build_result_tabs(self):
        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs, stretch=1)

        # Tab 1: 处理结果
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setFont(QFont("Microsoft YaHei", 11))
        self.tabs.addTab(self.result_text, "处理结果")

        # Tab 2: 商品明细
        self.detail_table = QTableWidget()
        self.detail_table.setColumnCount(4)
        self.detail_table.setHorizontalHeaderLabels(["商品名称", "规格", "单位", "数量"])
        self.detail_table.setAlternatingRowColors(True)
        self.detail_table.setSortingEnabled(True)
        self.detail_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.detail_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.detail_table.horizontalHeader().setStretchLastSection(False)
        self.detail_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.detail_table.setColumnWidth(0, 180)
        self.detail_table.setColumnWidth(1, 140)
        self.detail_table.setColumnWidth(2, 60)
        self.detail_table.setColumnWidth(3, 70)
        self.tabs.addTab(self.detail_table, "商品明细")

    def _build_done_buttons(self):
        self.done_container = QVBoxLayout()
        self.done_btn_layout = QHBoxLayout()
        self.done_btn_layout.setSpacing(12)

        self.close_btn = QPushButton("关闭")
        self.close_btn.setObjectName("closeBtn")
        self.close_btn.setFixedHeight(48)
        self.close_btn.clicked.connect(self.close)

        self.open_btn = QPushButton("📄 打开并退出")
        self.open_btn.setObjectName("openBtn")
        self.open_btn.setFixedHeight(48)
        self.open_btn.clicked.connect(self._open_and_close)
        self.open_btn.hide()

        self.done_btn_layout.addStretch()
        self.done_btn_layout.addWidget(self.close_btn)
        self.done_btn_layout.addWidget(self.open_btn)
        self.done_btn_layout.addStretch()
        self.done_container.addLayout(self.done_btn_layout)
        self.main_layout.addLayout(self.done_container)

        # 初始隐藏
        self._hide_widget(self.close_btn)
        self._show_done_buttons(False)

    def _setup_statusbar(self):
        self.status_bar = QStatusBar()
        self.status_bar.setStyleSheet("color: #999; font-size: 9pt;")
        self.status_label_sb = QLabel("Ctrl+Enter 执行 | Esc 关闭")
        self.status_bar.addPermanentWidget(self.status_label_sb)
        self.setStatusBar(self.status_bar)

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+Return"), self, self._execute)
        QShortcut(QKeySequence("Ctrl+Enter"), self, self._execute)
        QShortcut(QKeySequence("Escape"), self, self.close)
        QShortcut(QKeySequence("Ctrl+Q"), self, self.close)

    # ── 美化辅助 ────────────────────────────────────────────
    def _add_shadow(self, widget):
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(12)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(0, 0, 0, 30))
        widget.setGraphicsEffect(shadow)

    def _hide_widget(self, w):
        w.hide()

    def _show_widget(self, w):
        w.show()

    def _show_done_buttons(self, visible):
        if visible:
            self.close_btn.show()
        else:
            self.close_btn.hide()
            self.open_btn.hide()

    # ── 主题切换 ────────────────────────────────────────────
    def _apply_qss(self):
        self.setStyleSheet(DARK_QSS if self.is_dark else LIGHT_QSS)

    def _toggle_theme(self):
        self.is_dark = not self.is_dark
        self._apply_qss()
        self.theme_btn.setText("☀" if self.is_dark else "☽")
        self.theme_btn.setToolTip("切换亮色模式" if self.is_dark else "切换暗色模式")

        # Fusion 在暗色模式下表现更好
        if self.is_dark:
            QApplication.setStyle("Fusion")
            p = QApplication.palette()
            p.setColor(QPalette.Window, QColor(43, 43, 43))
            p.setColor(QPalette.WindowText, QColor(224, 224, 224))
            p.setColor(QPalette.Base, QColor(43, 43, 43))
            p.setColor(QPalette.AlternateBase, QColor(51, 51, 51))
            p.setColor(QPalette.Text, QColor(224, 224, 224))
            p.setColor(QPalette.Button, QColor(60, 60, 60))
            p.setColor(QPalette.ButtonText, QColor(224, 224, 224))
            p.setColor(QPalette.Highlight, QColor(40, 167, 69))
            p.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
            QApplication.setPalette(p)
        else:
            self._apply_light_palette()

        # 已有 QSS 覆盖，同时刷新样式
        self.setStyleSheet(DARK_QSS if self.is_dark else LIGHT_QSS)

    # ── 执行逻辑 ────────────────────────────────────────────
    def _execute(self, force=False):
        if self.is_running:
            return
        self.is_running = True
        self._op_start_time = time.time()  # 记录操作开始时间

        # UI 状态切换
        self.execute_btn.setEnabled(False)
        self.execute_btn.setText("执行中...")
        self.delete_cb.setEnabled(False)
        self.stats_cb.setEnabled(False)
        self._show_done_buttons(False)
        self.result_text.clear()
        self.detail_table.setRowCount(0)
        self.tabs.setCurrentIndex(0)
        self.detail_stats_lines = []
        self.detail_stats_total = 0
        self.output_file_path = ""
        self.result_data = None      # 清掉上一次结果，避免残留误导

        # 找 DLL
        dll = find_engine()
        if not dll:
            QMessageBox.critical(self, "错误",
                "Rust引擎 (jinhua_engine.dll) 未找到。\n\n"
                "请将 Rust 引擎放在程序同目录。")
            self._reset_ui()
            return

        # 启动工作线程（force=True = 「仍然生成」：忽略致命列识别问题继续出文件）
        self.worker = EngineWorker(dll, self.delete_cb.isChecked(), self.stats_cb.isChecked(),
                                   force=force)
        self.worker.progress_updated.connect(self._on_progress)
        self.worker.stats_received.connect(self._on_stats)
        self.worker.engine_finished.connect(self._on_finished)
        self.worker.engine_error.connect(self._on_error)
        QApplication.processEvents()  # 刷新"引擎处理中..."到 UI
        self.worker.start()

    def _reset_ui(self):
        self.is_running = False
        self.execute_btn.show()           # 必须在 setEnabled 之前恢复可见
        self.execute_btn.setEnabled(True)
        self.execute_btn.setText("执行表格转换")
        self.delete_cb.setEnabled(True)
        self.stats_cb.setEnabled(True)

    def _on_progress(self, pct, label, pct_str):
        self.progress_bar.setValue(pct)
        self.progress_label.setText(label)
        self.pct_label.setText(pct_str)

    def _on_stats(self, total, lines):
        self.detail_stats_total = total
        self.detail_stats_lines = lines

    def _on_finished(self, result_data):
        self.result_data = result_data
        self.progress_bar.setValue(100)
        self.progress_label.setText("执行完毕")
        self.pct_label.setText("100%")

        # 显示结果
        self._display_result(result_data)

        # 切换按钮
        self.execute_btn.hide()
        self._show_done_buttons(True)
        if result_data.get("file"):
            self.output_file_path = os.path.join(r"D:\订单表格", result_data["file"])
            if os.path.exists(self.output_file_path):
                self.open_btn.show()


        # 写入操作日志
        try:
            duration = int((time.time() - self._op_start_time) * 1000)
        except Exception:
            duration = 0
        log_operation(
            op_time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            operator=os.getlogin(),
            status="success",
            orders=result_data.get("orders", 0),
            customers=result_data.get("customers", 0),
            products=result_data.get("products", 0),
            matched=result_data.get("matched", 0),
            output_file=self.output_file_path,
            stats_enabled=1 if self.stats_cb.isChecked() else 0,
            delete_enabled=1 if self.delete_cb.isChecked() else 0,
            duration_ms=duration,
        )
        self.is_running = False

    def _on_error(self, msg, data=None):
        data = data or {}
        self.result_data = data      # 失败也是本次结果，避免页签显示上一次的旧数据
        diag_txt = _diags_text(data)
        fatal = bool(data.get("fatal"))
        # 中断也要把诊断显示在结果页，便于对照排查
        try:
            self._display_result(data)
        except Exception:
            pass
        # 诊断写入操作日志（截断，避免撑爆字段）
        try:
            duration = int((time.time() - self._op_start_time) * 1000)
        except Exception:
            duration = 0
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        op_id = log_operation(
            op_time=now_str,
            operator=os.getlogin(),
            status="error",
            orders=data.get("orders", 0),
            customers=data.get("customers", 0),
            products=data.get("products", 0),
            matched=data.get("matched", 0),
            output_file="",
            stats_enabled=1 if self.stats_cb.isChecked() else 0,
            delete_enabled=1 if self.delete_cb.isChecked() else 0,
            error_msg=(msg + " | " + diag_txt)[:2000],
            duration_ms=duration,
        )
        log_diagnostics(op_id, data.get("errors") or [], data.get("warnings") or [], now_str)
        self._reset_ui()

        box = QMessageBox(self)
        box.setWindowTitle("执行中止" if fatal else "错误")
        box.setIcon(QMessageBox.Critical if fatal else QMessageBox.Warning)
        box.setText(msg)
        if fatal:
            box.setInformativeText(
                "已中止，未生成输出文件（避免错误数据被导入优路达）。\n"
                "下面是定位到的具体问题；确认可以接受时也可选择「仍然生成」继续出文件。")
        box.setDetailedText(diag_txt)
        force_btn = box.addButton("仍然生成", QMessageBox.AcceptRole) if fatal else None
        box.addButton("取消", QMessageBox.RejectRole)
        box.exec()
        if force_btn is not None and box.clickedButton() is force_btn:
            self._execute(force=True)

    # ── 结果显示 ────────────────────────────────────────────
    def _display_result(self, d):
        fname = d.get("file", "")
        orders = d.get("orders", 0)
        customers = d.get("customers", 0)
        matched = d.get("matched", 0)
        products = d.get("products", 0)
        no_coords = d.get("no_coords", [])
        err = d.get("error")
        fatal = bool(d.get("fatal"))
        errors = d.get("errors") or []
        warnings = d.get("warnings") or []

        # ── 诊断区块：红=致命问题、黄=警告，都带出错字段与定位信息 ──
        diag_html = ""
        if errors:
            diag_html += ("<br><p style='color:#c0392b;font-weight:bold;'>"
                          "❌ 致命问题（列识别 / 数据校验）</p><ul style='color:#c0392b;'>")
            for e in errors:
                diag_html += "<li><b>%s</b>：%s" % (e.get("field", ""), e.get("problem", ""))
                if e.get("detail"):
                    diag_html += "<br><span style='font-size:9pt;'>%s</span>" % e.get("detail")
                diag_html += "</li>"
            diag_html += "</ul>"
        if warnings:
            diag_html += ("<p style='color:#b9770e;font-weight:bold;'>"
                          "⚠ 警告（已继续处理，建议核对）</p><ul style='color:#b9770e;'>")
            for w in warnings:
                diag_html += "<li><b>%s</b>：%s" % (w.get("field", ""), w.get("problem", ""))
                if w.get("detail"):
                    diag_html += "<br><span style='font-size:9pt;'>%s</span>" % w.get("detail")
                diag_html += "</li>"
            diag_html += "</ul>"

        if err:
            html = "<h3 style='color:red;'>执行中止</h3>"
            html += "<p><b>原因:</b> %s</p>" % err
            if fatal:
                html += "<p>本次<b>未生成</b>输出文件。</p>"
            html += "<p style='color:#555;'>处理订单数: %s 单；匹配明细: %s 单</p>" % (orders, matched)
        else:
            ok = (matched == orders)
            status_style = "" if ok else " style='color:red;'"
            status_text = "处理完成！" if ok else "处理未正确完成"
            html = f"<h3{status_style}>{status_text}</h3>"
            html += f"<p><b>生成文件:</b> {fname}</p>"
            html += f"<hr>"
            html += f"<p>处理订单数: <b>{orders}</b> 单</p>"
            if not ok:
                html += f"<p style='color:red;font-weight:bold;'>匹配明细: {matched} 单（{orders - matched} 单缺失）⚠</p>"
            else:
                html += f"<p>匹配明细: {matched} 单</p>"
            html += f"<p>合并后客户: <b>{customers}</b> 个</p>"
            if products > 0:
                html += f"<p>出库商品总数: <b>{products}</b></p>"

        html += diag_html

        if no_coords:
            html += f"<br><p style='color:red;'><b>⚠ 未匹配经纬度的客户 ({len(no_coords)} 个):</b></p>"
            html += "<ul style='color:red;'>"
            for name in no_coords[:30]:
                html += f"<li>{name}</li>"
            if len(no_coords) > 30:
                html += f"<li>...等共{len(no_coords)}个</li>"
            html += "</ul>"

        self.result_text.setHtml(html)

        # 摘要一行：引导用户切换到「商品明细」页签（与页签用同一数据源）
        rows = self._detail_rows()
        if rows:
            try:
                total_qty = sum(float(r[3]) for r in rows)
            except Exception:
                total_qty = 0
            self.result_text.append(
                f"\n  商品明细共 {len(rows)} 种（出库 {_numstr(total_qty)}），详情见「商品明细」页签\n"
            )
        else:
            self.result_text.append(
                "\n  （未勾选统计 或 D:\\明细表格 目录下无 .xls 文件）\n"
            )

        # 填充商品明细表格
        self._populate_detail_table()

    def _detail_rows(self):
        """商品明细行：(名称, 规格, 单位, 数量)。

        优先用汇总模块读到的**结构化数据**：规格里有没有空格、下划线、多长
        都不影响分列；只有拿不到结构化数据时，才退回解析引擎的等宽统计行。"""
        data = (self.result_data or {}).get("_summary")
        rows = []
        if data and not data.get("error"):
            for (name, spec, unit), v in data.get("items", []):
                rows.append((name, spec, unit, _numstr(sum(v["cust"].values()))))
            return rows
        for line in (self.detail_stats_lines or []):
            line = line.strip()
            if not line:
                continue
            rows.append(self._parse_stats_line(line))
        return rows

    def _populate_detail_table(self):
        self.detail_table.setSortingEnabled(False)
        self.detail_table.setRowCount(0)

        rows_data = self._detail_rows()
        if not rows_data:
            self.detail_table.setRowCount(1)
            hint = QTableWidgetItem("暂无数据（可能未勾选统计，或 D:\\明细表格 下无 .xls 文件）")
            hint.setTextAlignment(Qt.AlignCenter)
            hint.setFlags(Qt.NoItemFlags)
            self.detail_table.setItem(0, 0, hint)
            self.detail_table.setSpan(0, 0, 1, 4)
            return

        self.detail_table.setRowCount(len(rows_data))
        for row, (name, spec, unit, qty) in enumerate(rows_data):
            items = [
                QTableWidgetItem(name),
                QTableWidgetItem(spec),
                QTableWidgetItem(unit),
                QTableWidgetItem(qty),
            ]
            for col, item in enumerate(items):
                item.setTextAlignment(Qt.AlignLeft if col < 3 else Qt.AlignRight | Qt.AlignVCenter)
                self.detail_table.setItem(row, col, item)

        self.detail_table.setSortingEnabled(True)
        self.detail_table.resizeColumnsToContents()
        # 数量列固定窄宽，不被内容撑大
        self.detail_table.setColumnWidth(3, 70)

    @staticmethod
    def _parse_stats_line(line):
        """【仅回退用，切勿参考其切列方式】把引擎的等宽统计行切成四列。

        它是按空格倒着切（rsplit）的，规格里只要含空格就会串列——历史上
        “规格跑到单位列” 就是这么来的。正常路径请用 _detail_rows() 的结构化
        数据，这里只在拿不到结构化数据时兜底。
        """
        # 去掉前导两个空格
        s = line[2:] if len(line) > 2 and line[:2] == "  " else line

        # 从右向左解析：数量 和 单位 是两个独立字段
        parts = s.rstrip().rsplit(maxsplit=2)
        if len(parts) >= 3:
            rest, unit, qty = parts[0], parts[1], parts[2]
            # rest 包含 name + spec，spec 靠右对齐在 36 显示宽度处
            # 简单策略：rest 的最后一个词是 spec，其余是 name
            rest_parts = rest.rsplit(maxsplit=1)
            if len(rest_parts) >= 2:
                name, spec = rest_parts[0], rest_parts[1]
            else:
                name, spec = rest, ""
            return name.strip(), spec.strip(), unit.strip(), qty.strip()
        elif len(parts) == 2:
            return parts[0].strip(), "", parts[1].strip(), "0"
        return line.strip(), "", "", "0"

    # ── 打开/关闭 ────────────────────────────────────────────
    def _open_and_close(self):
        detail_dir = r'D:\订单表格\每日详情'
        os.makedirs(detail_dir, exist_ok=True)
        today = datetime.now().strftime('%m%d')
        txt_path = os.path.join(detail_dir, f'今日详情{today}.txt')
        txt_content = self.result_text.toPlainText()
        # 商品明细统计：直接用结构化数据自己排版（不再依赖引擎的等宽文本反解析）
        rows = self._detail_rows()
        if rows:
            w_name, w_spec, w_unit, w_qty = STATS_W
            total_qty = 0.0
            for *_, q in rows:
                try:
                    total_qty += float(q)
                except Exception:
                    pass
            txt_content += ("\n── 商品明细统计（共 %d 种，出库总计 %s）──\n"
                            % (len(rows), _numstr(total_qty)))
            txt_content += (f"  {_pad('商品名称', w_name)}{_pad('规格', w_spec)}"
                            f"{_pad('单位', w_unit)}数量\n")
            txt_content += "  " + "─" * 30 + "\n"
            for name, spec, unit, q in rows:
                txt_content += (f"  {_pad(name, w_name)}{_pad(spec, w_spec)}"
                                f"{_pad(unit, w_unit)}{_pad(q, w_qty)}\n")
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(txt_content)

        # 同时生成手机友好的商品汇总 HTML（与 TXT 同一天一份；失败不影响 TXT 与主流程）
        try:
            import summary
            data = (self.result_data or {}).get("_summary")
            if data and not data.get("error"):
                if summary.render(data, out_dir=detail_dir, quiet=True):
                    self.status_label_sb.setText(
                        "已生成：今日详情%s.txt ／ 商品汇总%s.html" % (today, today))
        except Exception:
            pass

        if self.output_file_path and os.path.exists(self.output_file_path):
            os.startfile(self.output_file_path)
        self.close()

    def closeEvent(self, event):
        if self.is_running:
            reply = QMessageBox.question(
                self, "确认", "程序正在运行，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                event.ignore()
                return
        event.accept()



# ============================================================================
# 启动动画 Splash
# ============================================================================
def show_splash():
    gif_path = os.path.join(BASE_DIR, "MLSX.gif")
    if not os.path.exists(gif_path):
        return None

    # 用 QLabel 包 QMovie 作为 Splash
    splash_pix = QPixmap(400, 300)
    splash_pix.fill(QColor(0, 0, 0, 0))

    splash = QSplashScreen(splash_pix)
    splash.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
    splash.setAttribute(Qt.WA_TranslucentBackground, True)

    # GIF 动画标签
    movie_label = QLabel(splash)
    movie = QMovie(gif_path)
    movie_label.setMovie(movie)
    movie_label.setAlignment(Qt.AlignCenter)
    movie_label.setGeometry(0, 0, 400, 300)
    movie.start()

    splash.show()
    return splash, movie

# ============================================================================
# 入口
# ============================================================================
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # 启动画面
    splash_tuple = show_splash()
    splash = splash_tuple[0] if splash_tuple else None

    # 模拟加载
    if splash:
        splash.showMessage("正在加载引擎...", Qt.AlignBottom | Qt.AlignCenter, QColor(100,100,100))
        QApplication.processEvents()
        time.sleep(0.5)

    # 创建主窗口
    window = JinhuaPyQtApp()

    if splash:
        # 淡入效果（WSLg 不支持，跳过）
        if os.environ.get("WSL_DISTRO_NAME"):
            splash.finish(window)
            window.show()
        else:
            window.setWindowOpacity(0.0)
            window.show()

            anim = QPropertyAnimation(window, b"windowOpacity")
            anim.setDuration(400)
            anim.setStartValue(0.0)
            anim.setEndValue(1.0)
            anim.setEasingCurve(QEasingCurve.OutCubic)

            def finish_splash():
                splash.finish(window)
            anim.finished.connect(finish_splash)
            anim.start()
    else:
        window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
