# D2Y-pyqt — PySide6 版（v3.1.2，发布主线）

与 `D2Y/` 的 ttkbootstrap 版**共用同一份 Rust 引擎 DLL**（`jinhua_engine.dll`，源码在 `..\D2Y\Rust\src\engine.rs`），GUI 层用 PySide6 重写。
最初是对比练习项目，现已转正为发布主线。详细设计见 `开发文档.md`。

- 最新安装包：[GitHub Releases v3.1.2](https://github.com/hotakun/jinhua-table-tool/releases/tag/v3.1.2)
- 本地安装包：`installer/金华聚火表格处理_v3.1.2_Setup.exe`

## 文件结构

```
D2Y-pyqt/
├── main.py                          # PySide6 GUI 主程序（单文件，1155 行）
├── jinhua_engine.dll                # Rust 引擎（复制自 ../D2Y/Rust/target/release/）
├── favicon.ico / MLSX.gif           # 程序图标 / 启动动画
├── JinhuaH.spec / build_exe.bat     # PyInstaller 打包配置与脚本
├── setup.iss / Default.isl          # Inno Setup 安装包脚本 + 中文语言文件
├── 开发文档.md / README.md          # 开发文档 / 本文件
├── tutorial/                        # PySide6 学习示例
├── dist/JinhuaH/                    # PyInstaller 输出（JinhuaH.exe + _internal/）
└── installer/                       # 安装包输出
```

## 运行方式

```powershell
cd D:\WFR\D2Y-pyqt
pip install PySide6
python main.py
```

## ttkbootstrap vs PySide6 对比

| 维度 | ttkbootstrap (D2Y/) | PySide6 (D2Y-pyqt/) |
|------|---------------------|----------------------|
| **框架** | tkinter + ttkbootstrap cosmo | PySide6 (Qt for Python) |
| **代码行数** | 496 行（hybrid.py） | 1155 行（main.py，含 QSS） |
| **启动动画** | 无 | MLSX.gif Splash Screen |
| **主题** | cosmo（20+ 预设改一个字符串） | Fusion + 自定义 QSS + 暗色模式切换 |
| **进度条** | ttk.Progressbar bootstyle="success-striped" | QProgressBar + QSS 圆角绿色条纹 |
| **结果展示** | Text 手拼字符串 + tag 上色 | QTextEdit HTML + QTableWidget 表格 |
| **商品明细** | Text 插入等宽文本 + 4 色 tag | QTableWidget（排序/交替色/复制） |
| **布局** | pack() | QVBoxLayout / QHBoxLayout 弹性布局 |
| **多线程** | threading.Thread + root.after | QThread + Signal/Slot |
| **暗色模式** | 不支持 | 一键切换（☽/☀ 按钮） |
| **快捷键** | 无 | Ctrl+Enter 执行 / Esc 关闭 |
| **卡片阴影** | 无 | QGraphicsDropShadowEffect |
| **按钮圆角** | ttk 方角 | QSS border-radius: 8px |
| **关闭确认** | Messagebox.yesno | QMessageBox.question |
| **窗口可缩放** | 固定 620×880 | 最小 580×680，可拉伸 |
| **表格功能** | 无 | 点击表头排序、Ctrl+C 复制、列宽拖拽 |
| **操作日志** | 无 | SQLite（D:\订单表格\logs\operations.db） |
| **自动更新提示** | 有（读 releases/latest） | 无（v3.1.x 已移除该代码） |
| **许可证** | MIT | LGPL（PySide6） |

## 功能清单

- [x] 启动 Splash 动画（MLSX.gif）
- [x] 窗口淡入过渡效果
- [x] 暗色/亮色主题切换
- [x] 处理结果 HTML 展示（标红未匹配行）
- [x] 商品明细 QTableWidget（排序/交替色）
- [x] QTabWidget 分 Tab 展示
- [x] QThread + 信号槽对接 Rust DLL
- [x] 进度条平滑更新
- [x] Ctrl+Enter 执行 / Esc 关闭快捷键
- [ ] GitHub 自动更新检测（**PySide6 版未实现**；tkinter v2.0.1 仍带此项，见开发文档第二节）
- [ ] 标题双击手动检查更新（同上，未实现）
- [x] SQLite 操作日志
- [x] TXT 导出 + os.startfile 打开
- [x] 关闭确认对话框
- [x] 卡片阴影美化
- [x] 打包发布（JinhuaH.spec + setup.iss，AppId 与旧版一致无缝升级）

## 刻意保留的差异

- 窗口默认 660×820（比原版稍矮，双 Tab 节省版面）
- 商品明细用表格而非纯文本（PySide6 优势项）
- 引擎层不做任何 GUI 相关处理，两个 GUI 共用同一 DLL

## 版本要点

- **v3.1.2**（已发布）：修复「外部订单号」列抢占订单号列导致的 **匹配明细 0 单 / 输出模板订单号列空白**（根因与影响面见 `开发文档.md` 第四节"销售订单列识别"）
