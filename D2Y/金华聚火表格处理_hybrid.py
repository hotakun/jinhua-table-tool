#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
金华聚火表格处理 - Python+ Rust 融合版
Python负责GUI (cosmo), Rust负责所有Excel处理
无需pandas, 启动极速
"""
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
import threading
import ctypes
import json
import urllib.request
import urllib.error
import ssl
import time
import os
import sys
import shutil
import random
from datetime import datetime, timedelta

CURRENT_VERSION = "2.0.1"
UPDATE_API = "https://api.github.com/repos/hotakun/jinhua-table-tool/releases/latest"
UPDATE_DL  = "https://github.com/hotakun/jinhua-table-tool/releases/download"


_ssl_ctx = None

def _fetch(url, timeout=10):
    """带 SSL 容错的 URL 请求"""
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


def _ver_newer(a: str, b: str) -> bool:
    """比较 a > b (去掉v前缀，按数字比较)"""
    def to_ints(v):
        return tuple(int(x) for x in v.lstrip("v").split(".") if x.isdigit())
    return to_ints(a) > to_ints(b)

# Rust引擎路径
ENGINE_EXE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "engine", "jinhua_engine.exe")


def find_engine():
    """查找Rust引擎DLL"""
    base = os.path.dirname(os.path.abspath(__file__))
    paths = [
        os.path.join(base, "jinhua_engine.dll"),
        os.path.join(base, "_internal", "jinhua_engine.dll"),
        os.path.join(base, "..", "Rust", "target", "release", "jinhua_engine.dll"),
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def check_for_update(root):
    """后台检查更新，支持自动下载"""
    try:
        with _fetch(UPDATE_API) as resp:
            data = json.loads(resp.read().decode())
        latest = data.get("tag_name", "").lstrip("v")
        if not latest or not _ver_newer(latest, CURRENT_VERSION):
            return
        # 找安装包下载链接
        dl_url = ""
        for asset in data.get("assets", []):
            name = asset.get("name", "")
            if name.endswith(".exe") and "Setup" in name:
                dl_url = asset["browser_download_url"]
                break
        if not dl_url:
            dl_url = f"{UPDATE_DL}/{data['tag_name']}/{data['tag_name']}_Setup.exe"

        root.after(0, lambda: _show_update_dialog(root, latest, dl_url))
    except:
        pass


def _show_update_dialog(root, latest, dl_url):
    if Messagebox.yesno(
        f"当前版本: v{CURRENT_VERSION}\n最新版本: v{latest}\n\n是否立即更新？",
        "发现新版本"
    ) == "Yes":
        threading.Thread(target=_download_and_install, args=(root, dl_url, latest), daemon=True).start()


def _download_and_install(root, url, version):
    import tempfile, tkinter.simpledialog
    save_path = os.path.join(tempfile.gettempdir(), f"JinhuaJuhuo_Setup_v{version}.exe")
    try:
        with _fetch(url, timeout=60) as resp:
            total = int(resp.headers.get("content-length", 0))
            downloaded = 0
            with open(save_path, 'wb') as f:
                while True:
                    chunk = resp.read(8192)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total > 0:
                        pct = int(downloaded * 100 / total)
                        root.after(0, lambda p=pct: setattr(root, '_dl_pct', p))
        root.after(0, lambda: _launch_installer(save_path))
    except Exception as e:
        root.after(0, lambda: Messagebox.show_error(f"下载失败: {str(e)[:200]}", "错误"))


def _launch_installer(path):
    if os.path.exists(path):
        os.startfile(path)
        os._exit(0)


# ============================================================================
# GUI
# ============================================================================

class PrintRedirector:
    """将stdout重定向到Text控件"""
    def __init__(self, text_widget):
        self.text = text_widget
    def write(self, s):
        self.text.after(0, self._write, s)
    def _write(self, s):
        self.text.config(state=tk.NORMAL)
        self.text.insert(tk.END, s)
        self.text.see(tk.END)
    def flush(self):
        pass


def _dwidth(s):
    """显示宽度: 中文≈4, ASCII≈2"""
    w = 0
    for c in s:
        o = ord(c)
        if o in (0x200B, 0x200C, 0x200D, 0xFEFF):
            continue
        if o in (0xFF08, 0xFF09):
            w += 1
        else:
            w += 4 if o > 0x2E80 else 2
    return w


def _pad(s, target_w):
    cur = _dwidth(s)
    return s + ' ' * max(0, target_w - cur)


class JinhuaHybridApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"金华聚火表格处理 v{CURRENT_VERSION}")
        self.root.geometry("620x880")

        for p in [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), 'favicon.ico'),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), '_internal', 'favicon.ico'),
            r'D:\WFR\D2Y\favicon.ico'
        ]:
            if os.path.exists(p): self.root.iconbitmap(default=p); break

        if sys.platform == "win32":
            try:
                import ctypes
                ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
            except: pass

        self.is_running = False
        self.result_message = ""
        self.output_file_path = ""
        self.detail_stats_lines = None
        self.detail_stats_widths = None
        self.detail_stats_total = 0
        self.detail_stats_raw = []
        self.delete_files_var = tk.BooleanVar(value=True)
        self.show_detail_stats_var = tk.BooleanVar(value=True)

        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.after(2000, lambda: threading.Thread(
            target=check_for_update, args=(self.root,), daemon=True).start())

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=(25, 15))
        main_frame.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure("TButton", font=("Microsoft YaHei", 14, "bold"), padding=(30, 12))
        style.configure("success.TButton", font=("Microsoft YaHei", 14, "bold"), padding=(30, 12))
        style.configure("danger.TButton", font=("Microsoft YaHei", 14, "bold"), padding=(30, 12))
        style.configure("secondary.TButton", font=("Microsoft YaHei", 14, "bold"), padding=(30, 12))

        ttk.Label(main_frame, text=f"金华聚火表格处理 v{CURRENT_VERSION}",
                  font=("Microsoft YaHei", 20, "bold"), cursor="hand2").pack(pady=(0, 2))
        ttk.Label(main_frame, text="订小易_转_优路达_表格处理",
                  font=("Microsoft YaHei", 10)).pack(pady=(0, 12))

        info_frame = ttk.LabelFrame(main_frame, text="使用说明")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(info_frame, wraplength=480, font=("Microsoft YaHei", 9),
                  text="  1、D:\\订单表格\\销售订单.xls\n  2、D:\\订单表格\\客户经纬度.xls\n  3、D:\\订单表格\\优路达导入模板.xlsx\n  4、D:\\明细表格\\销售订单详细 - *.xls"
                  ).pack(anchor="w")

        chk_frame = ttk.Frame(main_frame)
        chk_frame.pack(fill=tk.X, pady=(0, 10))
        self.delete_checkbox = ttk.Checkbutton(chk_frame,
            text="执行完毕后删除 D:\\明细表格 目录", variable=self.delete_files_var)
        self.delete_checkbox.pack(anchor="w")
        self.detail_stats_checkbox = ttk.Checkbutton(chk_frame,
            text="查看商品明细数量", variable=self.show_detail_stats_var)
        self.detail_stats_checkbox.pack(anchor="w")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(0, 15))
        self.execute_button = ttk.Button(btn_frame, text="执行表格转换",
            bootstyle=SUCCESS, command=self.execute_conversion)
        self.execute_button.pack()

        prog_frame = ttk.LabelFrame(main_frame, text="进度")
        prog_frame.pack(fill=tk.X, pady=(0, 10))
        self.progress_label = ttk.Label(prog_frame, text="就绪", font=("Microsoft YaHei", 11))
        self.progress_label.pack(pady=(3, 3))
        self.progress_bar = ttk.Progressbar(prog_frame, length=420, mode=DETERMINATE,
                                             bootstyle="success-striped")
        self.progress_bar.pack(pady=(0, 3))
        self.percentage_label = ttk.Label(prog_frame, text="0%", font=("Microsoft YaHei", 10))
        self.percentage_label.pack(pady=(0, 3))

        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.result_text = tk.Text(result_frame, font=("Microsoft YaHei", 11), height=22, wrap=tk.NONE)
        self.result_text.tag_config("red", foreground="#e74c3c", font=("Microsoft YaHei", 11, "bold"))
        self.result_text.tag_config("head", font=("Microsoft YaHei", 11, "bold"))
        self.result_text.tag_config("stripe", background="#f5f5f5")
        self.result_text.bind("<Key>", lambda e: "break")
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = ttk.Scrollbar(result_frame, orient=tk.VERTICAL,
                                command=self.result_text.yview, bootstyle="round")
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.configure(yscrollcommand=scroll.set)

        self.status_label = ttk.Label(main_frame, font=("Microsoft YaHei", 10))
        self.status_label.pack_forget()

    def execute_conversion(self):
        if self.is_running:
            return
        self.is_running = True
        self.execute_button.config(state=tk.DISABLED, text="执行中...", bootstyle=SECONDARY)
        self.delete_checkbox.config(state=tk.DISABLED)
        self.detail_stats_checkbox.config(state=tk.DISABLED)
        self.output_file_path = ""
        if hasattr(self, '_btn_row'):
            self._btn_row.destroy()
        self.execute_button.pack()
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)

        self._old_stdout = sys.stdout
        sys.stdout = PrintRedirector(self.result_text)

        thread = threading.Thread(target=self.run_task, daemon=True)
        thread.start()

    def run_task(self):
        try:
            self.update_progress(0, "正在启动引擎...", "0%")
            engine = find_engine()

            if engine:
                # Rust引擎模式
                self._run_rust_engine(engine)
            else:
                # 降级到Python模式
                self._run_fallback()
        except Exception as e:
            self.root.after(0, lambda: self.show_error(f"执行失败: {str(e)}"))

    def _run_rust_engine(self, dll_path):
        try:
            self.update_progress(5, "检查并创建目录...", "5%")
            for d in [r'D:\订单表格', r'D:\明细表格', r'D:\明细表格\MXBG2']:
                os.makedirs(d, exist_ok=True)

            lib = ctypes.CDLL(dll_path)

            # 进度回调
            CB = ctypes.CFUNCTYPE(None, ctypes.c_uint32, ctypes.c_char_p)
            @CB
            def progress_cb(pct, label_ptr):
                try:
                    label = label_ptr.decode('utf-8') if label_ptr else ""
                    if label.startswith("STATS:"):
                        # stats:STATS:total|"line1","line2",...
                        parts = label[6:].split("|", 1)
                        self.detail_stats_total = int(parts[0])
                        self.detail_stats_lines = json.loads(f"[{parts[1]}]") if len(parts) > 1 else []
                    else:
                        self.update_progress(pct, label, f"{pct}%")
                except: pass

            lib.engine_process.argtypes = [ctypes.c_uint8, ctypes.c_uint8, CB]
            lib.engine_process.restype = ctypes.c_void_p
            lib.engine_free.argtypes = [ctypes.c_void_p]

            # 后台线程调用 DLL，避免阻塞 GUI
            def call_dll():
                self.update_progress(5, "引擎处理中...", "5%")
                delete = 1 if self.delete_files_var.get() else 0
                stats = 1 if self.show_detail_stats_var.get() else 0
                ptr = lib.engine_process(delete, stats, progress_cb)
                result_json = ctypes.cast(ptr, ctypes.c_char_p).value.decode('utf-8')
                lib.engine_free(ptr)
                return result_json

            # 在线程中调用（ctypes 会阻塞当前线程，不影响 GUI）
            import concurrent.futures
            future = concurrent.futures.ThreadPoolExecutor(1).submit(call_dll)
            while not future.done():
                time.sleep(0.1)
                self.root.update()

            result_data = json.loads(future.result())
            if "error" in result_data:
                raise Exception(result_data["error"])

            self._build_result_message(result_data)
            self.update_progress(100, "执行完毕", "100%")
            time.sleep(0.3)
            self.root.after(0, self.show_completion)

        except OSError:
            Messagebox.show_warning("Rust引擎DLL未找到", "提示")
            self._run_fallback()
        except Exception as e:
            self.root.after(0, lambda: self.show_error(f"引擎异常: {str(e)}"))

    def _run_fallback(self):
        """Rust不可用时降级为模拟处理"""
        self.update_progress(5, "降级模式: 请使用原版Python工具", "5%")
        time.sleep(1)
        self.root.after(0, lambda: self.show_error(
            "Rust引擎 (jinhua_engine.exe) 未找到。\n\n"
            "请将Rust引擎放在以下任一位置:\n"
            "  • 程序同目录\n"
            "  • engine\\jinhua_engine.exe\n"
            "  • ..\\Rust\\target\\release\\jinhua_engine.exe\n\n"
            "或使用原版Python工具。"
        ))

    def _build_result_message(self, d):
        """用 Rust 返回的数据构建格式化结果文本"""
        sep = "─" * 40
        fname = d.get("file", "")
        orders = d.get("orders", 0)
        customers = d.get("customers", 0)
        matched = d.get("matched", 0)
        products = d.get("products", 0)
        no_coords = d.get("no_coords", [])

        msg = f"处理完成！\n生成文件: {fname}\n{sep}\n"
        msg += f"处理订单数: {orders} 单\n"
        if matched != orders:
            msg += "\x02"  # 标红
        msg += f"匹配明细: {matched} 单\n"
        msg += f"合并后客户: {customers} 个\n"
        if products > 0:
            msg += f"出库商品总数 {products}\n"
        if no_coords:
            msg += f"\n⚠ 未匹配经纬度的客户 ({len(no_coords)} 个):\n"
            for name in no_coords[:30]:
                msg += f"  • {name}\n"
            if len(no_coords) > 30:
                msg += f"  ...等共{len(no_coords)}个\n"

        self.result_message = msg
        self.output_file_path = os.path.join(r"D:\订单表格", fname) if fname else ""

    def update_progress(self, value, text, percentage):
        self.root.after(0, lambda: self._do_update(value, text, percentage))

    def _do_update(self, value, text, percentage):
        self.progress_bar['value'] = value
        self.progress_label.config(text=text)
        self.percentage_label.config(text=percentage)
        self.root.update()

    def show_completion(self):
        sys.stdout = getattr(self, '_old_stdout', sys.stdout)
        self.root.update()
        self.status_label.config(text="执行完毕")

        self.execute_button.pack_forget()
        self._btn_row = ttk.Frame(self.execute_button.master)
        self._btn_row.pack(pady=(0, 15))
        ttk.Button(self._btn_row, text="关闭", bootstyle=DANGER,
                   command=self.close_window).pack(side=tk.LEFT, padx=5)
        if self.output_file_path and os.path.exists(self.output_file_path):
            ttk.Button(self._btn_row, text="📄 打开并退出", bootstyle=SUCCESS,
                       command=self._open_and_close).pack(side=tk.LEFT, padx=5)

        self.is_running = False
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert("1.0", self.result_message)

        # 标红
        count = int(self.result_text.index('end-1c').split('.')[0])
        for i in range(1, count + 1):
            line_text = self.result_text.get(f"{i}.0", f"{i}.end")
            if "\x02" in line_text:
                cleaned = line_text.replace("\x02", "")
                self.result_text.delete(f"{i}.0", f"{i}.end")
                self.result_text.insert(f"{i}.0", cleaned)
                self.result_text.tag_add("red", f"{i}.0", f"{i}.end")
            elif line_text.startswith("⚠") or line_text.startswith("  •"):
                self.result_text.tag_add("red", f"{i}.0", f"{i}.end")
            elif line_text.startswith("──"):
                self.result_text.tag_add("head", f"{i}.0", f"{i}.end")

        # 商品明细
        if self.detail_stats_lines:
            w_name, w_spec, w_unit = self.detail_stats_widths or (46, 36, 10)
            self.result_text.insert(tk.END, f"\n── 商品明细统计（共 {len(self.detail_stats_lines)} 种，出库总计 {self.detail_stats_total}）──\n")
            self.result_text.insert(tk.END, f"  {_pad('商品名称', w_name)}{_pad('规格', w_spec)}{_pad('单位', w_unit)}{'数量':>6}\n")
            self.result_text.insert(tk.END, "  " + "─"*40 + "\n")
            colors = ["#f8f8f8", "#eef6ff", "#f5fff0", "#fff8f0"]
            prev_name = None
            color_idx = -1
            for line in self.detail_stats_lines:
                name = line.strip().split()[0] if line.strip() else ""
                if name != prev_name:
                    color_idx = (color_idx + 1) % len(colors)
                    prev_name = name
                tag = f"c{color_idx}"
                self.result_text.tag_config(tag, background=colors[color_idx])
                start = self.result_text.index(tk.END + "-1c")
                self.result_text.insert(tk.END, line + "\n")
                self.result_text.tag_add(tag, start, tk.END + "-1c")

    def _open_and_close(self):
        detail_dir = r'D:\订单表格\每日详情'
        os.makedirs(detail_dir, exist_ok=True)
        today = datetime.now().strftime('%m%d')
        with open(os.path.join(detail_dir, f'今日详情{today}.txt'), 'w', encoding='utf-8') as f:
            f.write(self.result_text.get("1.0", tk.END))
        if self.output_file_path and os.path.exists(self.output_file_path):
            os.startfile(self.output_file_path)
        self.close_window()

    def show_error(self, message):
        sys.stdout = getattr(self, '_old_stdout', sys.stdout)
        self.is_running = False
        if hasattr(self, '_btn_row'):
            self._btn_row.destroy()
        self.execute_button.pack()
        self.execute_button.config(state=tk.NORMAL, text="执行表格转换", bootstyle=SUCCESS)
        self.delete_checkbox.config(state=tk.NORMAL)
        self.detail_stats_checkbox.config(state=tk.NORMAL)
        Messagebox.show_error(message, "错误")

    def close_window(self):
        self.root.destroy()

    def on_closing(self):
        if self.is_running:
            if Messagebox.yesno("程序正在运行，确定要退出吗？", "确认") != "Yes":
                return
        self.root.destroy()


def main():
    root = ttk.Window(themename="cosmo")
    app = JinhuaHybridApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
