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
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QCheckBox, QGroupBox, QProgressBar,
    QTextEdit, QTableWidget, QTableWidgetItem, QTabWidget,
    QSplashScreen, QFrame, QToolButton, QStatusBar, QMessageBox,
    QHeaderView, QGraphicsDropShadowEffect,
)
from PySide6.QtCore import (
    Qt, QPoint, QThread, Signal, QTimer, QPropertyAnimation, QEasingCurve,
)
from PySide6.QtGui import (
    QFont, QIcon, QPixmap, QMovie, QColor, QPalette,
    QShortcut, QKeySequence,
)

# ============================================================================
# 常量
# ============================================================================
CURRENT_VERSION = "3.1.2"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 操作日志数据库
LOG_DB_DIR = r"D:\订单表格\logs"
LOG_DB_PATH = os.path.join(LOG_DB_DIR, "operations.db")

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
    conn.commit()
    conn.close()

def log_operation(op_time, operator, status, orders=0, customers=0,
                  products=0, matched=0, output_file="",
                  stats_enabled=0, delete_enabled=0,
                  error_msg="", duration_ms=0):
    """写入一条操作日志，失败不影响主流程"""
    try:
        os.makedirs(LOG_DB_DIR, exist_ok=True)
        conn = sqlite3.connect(LOG_DB_PATH)
        conn.execute("""
            INSERT INTO operations
                (op_time, operator, status, orders, customers,
                 products, matched, output_file, stats_enabled,
                 delete_enabled, error_msg, duration_ms)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (op_time, operator, status, orders, customers,
              products, matched, output_file, stats_enabled,
              delete_enabled, error_msg, duration_ms))
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

def _pad(s, target_w):
    return s + ' ' * max(0, target_w - _dwidth(s))

# Rust 引擎统计列宽度: name=46, spec=36, unit=10, qty(右对齐)=6
STATS_W = (46, 36, 10, 6)

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
    engine_error = Signal(str)                  # error message

    def __init__(self, dll_path, delete_files, show_stats, parent=None):
        super().__init__(parent)
        self.dll_path = dll_path
        self.delete_files = delete_files
        self.show_stats = show_stats

    def run(self):
        try:
            for d in [r'D:\订单表格', r'D:\明细表格', r'D:\明细表格\MXBG2']:
                os.makedirs(d, exist_ok=True)

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

            lib.engine_process.argtypes = [ctypes.c_uint8, ctypes.c_uint8, CB]
            lib.engine_process.restype = ctypes.c_void_p
            lib.engine_free.argtypes = [ctypes.c_void_p]

            self.progress_updated.emit(5, "引擎处理中...", "5%")
            delete = 1 if self.delete_files else 0
            stats = 1 if self.show_stats else 0
            ptr = lib.engine_process(delete, stats, progress_cb)
            result_json = ctypes.cast(ptr, ctypes.c_char_p).value.decode('utf-8')
            lib.engine_free(ptr)

            result_data = json.loads(result_json)
            if "error" in result_data:
                self.engine_error.emit(result_data["error"])
            else:
                self.engine_finished.emit(result_data)

        except OSError as e:
            self.engine_error.emit(f"Rust引擎DLL加载失败: {str(e)}")
        except Exception as e:
            self.engine_error.emit(f"引擎异常: {str(e)}")


# ============================================================================
# 主窗口
# ============================================================================
class JinhuaPyQtApp(QMainWindow):
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

        self._setup_ui()
        self._setup_statusbar()
        self._setup_shortcuts()

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
        h_layout.addWidget(self.help_btn, alignment=Qt.AlignTop)
        h_layout.addWidget(self.theme_btn, alignment=Qt.AlignTop)
        self.main_layout.addLayout(h_layout)

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
    def _execute(self):
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

        # 找 DLL
        dll = find_engine()
        if not dll:
            QMessageBox.critical(self, "错误",
                "Rust引擎 (jinhua_engine.dll) 未找到。\n\n"
                "请将 Rust 引擎放在程序同目录。")
            self._reset_ui()
            return

        # 启动工作线程
        self.worker = EngineWorker(dll, self.delete_cb.isChecked(), self.stats_cb.isChecked())
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

    def _on_error(self, msg):
        # 记录错误日志
        try:
            duration = int((time.time() - self._op_start_time) * 1000)
        except Exception:
            duration = 0
        log_operation(
            op_time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            operator=os.getlogin(),
            status="error",
            orders=0,
            customers=0,
            products=0,
            matched=0,
            output_file="",
            stats_enabled=1 if self.stats_cb.isChecked() else 0,
            delete_enabled=1 if self.delete_cb.isChecked() else 0,
            error_msg=msg,
            duration_ms=duration,
        )
        self._reset_ui()
        QMessageBox.critical(self, "错误", msg)

    # ── 结果显示 ────────────────────────────────────────────
    def _display_result(self, d):
        fname = d.get("file", "")
        orders = d.get("orders", 0)
        customers = d.get("customers", 0)
        matched = d.get("matched", 0)
        products = d.get("products", 0)
        no_coords = d.get("no_coords", [])

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

        if no_coords:
            html += f"<br><p style='color:red;'><b>⚠ 未匹配经纬度的客户 ({len(no_coords)} 个):</b></p>"
            html += "<ul style='color:red;'>"
            for name in no_coords[:30]:
                html += f"<li>{name}</li>"
            if len(no_coords) > 30:
                html += f"<li>...等共{len(no_coords)}个</li>"
            html += "</ul>"

        self.result_text.setHtml(html)

        # 摘要一行：引导用户切换到「商品明细」页签
        if self.detail_stats_lines:
            self.result_text.append(
                f"\n  商品明细共 {sum(1 for l in self.detail_stats_lines if l.strip())} 种"
                f"（出库 {self.detail_stats_total}），详情见「商品明细」页签\n"
            )
        else:
            self.result_text.append(
                "\n  （未勾选统计 或 D:\\明细表格 目录下无 .xls 文件）\n"
            )

        # 填充商品明细表格
        self._populate_detail_table()

    def _populate_detail_table(self):
        self.detail_table.setSortingEnabled(False)
        self.detail_table.setRowCount(0)

        if not self.detail_stats_lines:
            self.detail_table.setRowCount(1)
            hint = QTableWidgetItem("暂无数据（可能未勾选统计，或 D:\\明细表格 下无 .xls 文件）")
            hint.setTextAlignment(Qt.AlignCenter)
            hint.setFlags(Qt.NoItemFlags)
            self.detail_table.setItem(0, 0, hint)
            self.detail_table.setSpan(0, 0, 1, 4)
            return

        rows_data = []
        for line in self.detail_stats_lines:
            line = line.strip()
            if not line:
                continue
            name, spec, unit, qty = self._parse_stats_line(line)
            rows_data.append((name, spec, unit, qty))

        if not rows_data:
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
        """解析 Rust 引擎预格式化的统计行: '  名称  规格  单位  数量'
           引擎使用 Chinese-width 对齐 (中文=4, ASCII=2)
           已知宽度: name=46, spec=36, unit=10, qty 右对齐 6
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
        if self.detail_stats_lines:
            w_name, w_spec, w_unit, w_qty = STATS_W
            nkinds = sum(1 for l in self.detail_stats_lines if l.strip())
            txt_content += f"\n── 商品明细统计（共 {nkinds} 种，出库总计 {self.detail_stats_total}）──\n"
            txt_content += f"  {_pad('商品名称', w_name)}{_pad('规格', w_spec)}{_pad('单位', w_unit)}{'数量':>{w_qty}}\n"
            txt_content += "  " + "─" * 30 + "\n"
            for line in self.detail_stats_lines:
                if line.strip():
                    txt_content += line + "\n"
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(txt_content)

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
