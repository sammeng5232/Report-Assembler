# Report Assembler

**保险行业季度调研报告生成器** — 多源 DOCX 合成 · 章节预览可调 · 格式模板复用 · 导出 DOCX / PDF

把各部门提交的分板块 Word 材料，自动识别章节、合并汇总，并套用统一版式，生成可交付的季度终版报告。

---

## 功能概览

| 能力 | 说明 |
|------|------|
| 多文件汇总 | 导入多份 DOCX，按章节编号合并为单一终稿 |
| 智能识别 | 支持「第 N 章」「1 / 2.1 / 2.1.1」等编号及 Word 标题样式 |
| 预览与修订 | 生成前可改标题、编号、层级，并调整章节顺序 |
| 格式模板 | 可选历史上季终版，对齐章节顺序、标题与字体 |
| 图文保留 | 表格、图片、加粗等随内容迁移 |
| 题注与引用 | 图/表题注统一编号，正文交叉引用尽量理顺 |
| 一键交付 | 输出 DOCX；本机有 Word / LibreOffice 时可导出 PDF |
| 坏文件隔离 | 损坏、加密、伪装 docx 单独记录，不拖垮整批 |

---

## 快速开始

### Windows（推荐）

1. 打开本目录，双击 **`保险报告生成器.exe`**
2. 添加本季各部门 DOCX（或使用下方样例）
3. （可选）启用格式模板，选择上季终版 docx
4. 点击 **预览章节** → 确认/调整 → **生成终版报告**

> 请保持 `保险报告生成器.exe` 与 `_internal` 文件夹在同一目录，不要只拷贝 exe。

### macOS

详见 [Mac使用说明.md](Mac使用说明.md)。

简要步骤：

```bash
cd 工具
chmod +x run_mac.command
./run_mac.command
```

首次会创建虚拟环境并安装依赖，随后启动与 Windows 相同的图形界面。  
也可在 Mac 上执行 `build_mac_app.sh` 打包为 `.app`（须在本机 macOS 上构建）。

### 从源码运行（Windows / macOS）

```bash
cd 工具
python -m venv .venv

# Windows
.venv\Scripts\activate
# macOS / Linux
# source .venv/bin/activate

pip install -r requirements.txt
# Windows 导出 PDF 还需要：pip install pywin32
# 并安装本机 Microsoft Word（或 LibreOffice 作为回退）

python desktop_app.py
```

**依赖：** Python 3.10+，`python-docx`、`lxml`；Windows PDF 导出建议 `pywin32` + Word。

---

## 使用流程

```
添加 DOCX 材料
    → （可选）加载上季终版作格式模板
    → 预览章节（识别结构，可改标题/编号/顺序）
    → 填写报告周期、标题、单位、页眉与字体
    → 生成终版 DOCX（及 PDF）
```

**典型场景：** 用 `保险行业季度调研报告_终版_2026Q2.docx` 作模板，导入 `输入文件_2026Q3/` 下六份材料，预览确认后生成第三季度终版。

---

## 目录结构

```
Report Assembler/
├── README.md                 # 本说明
├── Mac使用说明.md             # macOS 专用说明
├── 保险报告生成器.exe          # Windows 可执行程序
├── _internal/                # 打包运行时（勿删）
├── 工具/                     # 源码
│   ├── desktop_app.py        # 桌面界面入口
│   ├── report_aggregator.py  # 合成核心
│   ├── app_utils.py          # 跨平台工具
│   ├── requirements.txt
│   ├── run_mac.command       # Mac 一键启动
│   └── build_mac_app.sh      # Mac 打包 .app
├── 输入文件_2026Q2/           # Q2 样例输入
├── 输入文件_2026Q3/           # Q3 样例输入
├── 演示输出/                  # 演示生成结果
├── 宣传/                     # 演示视频、Beamer 文稿等
└── 保险行业季度调研报告_终版_*.docx / *.pdf
```

---

## 样例数据

| 路径 | 用途 |
|------|------|
| `输入文件_2026Q3/` | 本季六份分板块 DOCX |
| `保险行业季度调研报告_终版_2026Q2.docx` | 上季终版，可作格式模板 |
| `演示输出/` | 一键演示生成的终版示例 |

命令行也可直接合成（不启动界面）：

```bash
cd 工具
python report_aggregator.py <输入目录> <输出.docx> \
  --period "2026年第三季度" \
  --title "中国保险行业调研报告" \
  --org "保险行业研究中心" \
  --date "2026年10月"
```

---

## 输出与 PDF

- **DOCX：** 始终生成；含封面、目录域、页眉页脚、正文与题注。
- **PDF：**
  - **Windows：** 优先本机 Word（COM / pywin32）
  - **macOS：** 优先 Microsoft Word（AppleScript）
  - **回退：** LibreOffice (`soffice`)
  - 均不可用时仍保留 DOCX，可在 Word / WPS / Pages 中手动另存 PDF

打开 DOCX 后，建议在 Word 中 **全选 → 右键更新域（或 Ctrl+A → F9）**，以刷新目录、题注与页码。

---

## 输入材料建议

1. 每个文件使用清晰的章节编号（如 `1 概述`、`2.1 政策环境`、`2.1.1 …`）
2. 同编号多文件会合并内容；标题不一致时预览中会标为冲突，请人工核对
3. 请提交标准 `.docx`（勿用加密文档或把 PDF/旧版 .doc 改后缀）
4. 单个文件建议不超过 200MB

---

## 宣传与演示

- 操作演示视频：`宣传/保险报告生成器_操作演示.mp4`
- 介绍幻灯片：`宣传/保险报告生成器_Beamer宣传.pdf`

---

## 平台说明

| 方式 | Windows | macOS |
|------|---------|-------|
| 预编译程序 | `保险报告生成器.exe` + `_internal` | 需本机打包 `.app` 或源码运行 |
| 源码 GUI | `python desktop_app.py` | `run_mac.command` |
| PDF 自动导出 | Word 或 LibreOffice | Word 或 LibreOffice |

Windows 的 `.exe` **不能**在 Mac 上运行；Mac 的 `.app` 须在 macOS 上构建。

---

## 许可证与用途

面向保险行业季度调研报告的内部编务 / 研究使用。请勿将含敏感业务数据的样例报告对外扩散；分发软件时建议仅提供脱敏样例。
