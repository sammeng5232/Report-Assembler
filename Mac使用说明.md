# 保险报告生成器 — MacBook 使用说明

## 结论

| 方式 | 是否可行 | 说明 |
|------|----------|------|
| Windows 的 `.exe` | ❌ | 不能在 Mac 上运行 |
| 源码直接运行 | ✅ | 推荐：拷贝 `工具/` 到 Mac 后双击 `run_mac.command` |
| 打成 `.app` | ✅ | **必须在 Mac 上**用 `build_mac_app.sh` 打包（无法在 Windows 交叉编译） |

核心合成逻辑（python-docx）本身跨平台；此前卡点是「打开文件」和「Word 转 PDF」，现已改为按系统自动选择。

---

## 一、最快上手（源码运行）

1. 把整个 `报告合成` 文件夹拷到 Mac（U 盘 / 网盘 / AirDrop 均可）。
2. 安装 Python 3.10+  
   - [python.org](https://www.python.org/downloads/macos/) 官方安装包（自带 Tk），或  
   - `brew install python`（若缺 Tk：`brew install python-tk`）
3. 在访达中进入 `报告合成/工具/`，**双击** `run_mac.command`  
   - 若提示无法打开：右键 → 打开，或在终端执行：
     ```bash
     cd ~/Desktop/报告合成/工具   # 按实际路径改
     chmod +x run_mac.command
     ./run_mac.command
     ```
4. 首次会自动创建 `.venv` 并安装依赖，随后弹出与 Windows 相同的图形界面。

可选：安装 [Microsoft Word for Mac](https://www.microsoft.com/microsoft-365/word) 或 [LibreOffice](https://www.libreoffice.org/)，生成报告时才能自动导出 PDF；没有的话 **DOCX 仍可正常生成**，再用 Word/Pages 另存 PDF。

---

## 二、打成 Mac 应用（.app）

在 **MacBook 本机** 终端执行：

```bash
cd ~/Desktop/报告合成/工具
chmod +x build_mac_app.sh
./build_mac_app.sh
```

生成物：`工具/dist/保险报告生成器.app`  
可拖到「应用程序」。若 Gatekeeper 拦截：

```bash
xattr -cr ~/Desktop/报告合成/工具/dist/保险报告生成器.app
```

> **注意**：不能在 Windows 上替你打出可用的 `.app`（签名、二进制架构、系统库都不同）。请把源码带到 Mac 上打包。

---

## 三、PDF 导出策略（自动）

1. **macOS + 已装 Word** → AppleScript 调用 Word 导出  
2. **否则若装了 LibreOffice** → `soffice` 无界面转换  
3. **都没有** → 只出 DOCX，界面提示可手动另存 PDF  

Windows 仍优先用本机 Word COM（需 pywin32）。

---

## 四、字体说明

- Mac 界面默认：`PingFang SC`（苹方）  
- 报告正文默认：`Songti SC`（宋体-简）  
- 下拉框仍保留「微软雅黑 / 宋体」等，便于与 Windows 终版对齐；若本机没有该字体，Word 打开时会回退到相近字体。

---

## 五、与 Windows 共用材料

两边可共用：

- `输入文件_2026Q3/*.docx`
- 格式模板 `保险行业季度调研报告_终版_2026Q2.docx`
- 生成出的 `*.docx` 终版

仅不可共用：`保险报告生成器.exe` / `_internal`（Windows 专用）。
