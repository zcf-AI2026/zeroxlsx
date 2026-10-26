# zeroxlsx

**零依赖写真正的 `.xlsx` —— 样式、公式、合并单元格、列宽、内嵌浮动图片，全用标准库完成。**

[![CI](https://github.com/zcf-AI2026/zeroxlsx/actions/workflows/ci.yml/badge.svg)](https://github.com/zcf-AI2026/zeroxlsx/actions/workflows/ci.yml)
[![PyPI](https://img.shields.io/pypi/v/zeroxlsx.svg)](https://pypi.org/project/zeroxlsx/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

`pip install openpyxl` 会带进四个包，`XlsxWriter` 也有自己的依赖树。可有时候你只是想
**吐出一个表格文件** —— 在 Lambda 里、在 CI 里、在不能联网的机器上、在一个单文件脚本里
—— 为了这点事装一棵依赖树，不划算。

`zeroxlsx` 用手写 OOXML 解决这件事：只用 `zipfile` + `xml`，别的什么都不用。
**连图片都能内嵌，且不需要 Pillow** —— 宽高直接读 PNG/JPEG 文件头。

```python
from zeroxlsx import Workbook, Style

book = Workbook()
sheet = book.add_sheet("报表")

sheet.merge("A1:D1", "季度汇总", Style(bold=True, size=14))
for column, label in enumerate(["区域", "单数", "金额", "毛利率"], start=1):
    sheet.write(2, column, label, Style(bold=True, fill="#F2E8D5", h_align="center"))

sheet.write(3, 1, "北区")
sheet.write(3, 2, 128)
sheet.write(3, 3, 48230.50, Style(number_format="#,##0.00"))
sheet.write(3, 4, 0.312, Style(number_format="0.0%"))

sheet.add_image("logo.png", anchor="F2", width=72)   # PNG/JPEG，不需要 Pillow
book.save("报表.xlsx")
```

## 和现成方案比

| 能力 | openpyxl | XlsxWriter | **zeroxlsx** |
|---|---|---|---|
| 写 xlsx | ✅ | ✅ | ✅ |
| 读 xlsx | ✅ | ❌ | ❌（有意不做） |
| 图表 / 数据透视 | ✅ | ✅ | ❌ |
| **运行时依赖** | 4 个 | 2 个 | **0 个** |
| 免 Pillow 内嵌图片 | ❌ | ✅ | ✅ |
| 样式 / 合并 / 列宽 / 公式 | ✅ | ✅ | ✅ |
| 输出可确定性（同输入同字节） | ❌ | — | ✅ |

需要**读**文件或做图表，用 openpyxl / XlsxWriter —— 它们在这方面很强。
`zeroxlsx` 只负责“把文件写出来”这件事，代价是零安装。

## 安装

```bash
pip install zeroxlsx
```

Python 3.9+，无依赖，wheel 体积约 20 KB。

## 支持范围

- **值**：字符串（inline 写法，不建共享字符串表）、数字、布尔、日期/时间（写成真正的
  Excel 序列值）、公式（`"=SUM(A1:A3)"`）
- **样式**：粗体 / 斜体 / 字号 / 字色 / 纯色填充 / 自动换行 / 水平与垂直对齐 /
  数字格式（内置 id 或任意自定义格式串）；相同样式自动合并为一个 `cellXfs`
- **版式**：合并区域、列宽、多工作表（按 Excel 规则校验名字：≤31 字符、禁 `[]:*?/\`、不重名）
- **图片**：PNG / JPEG，锚定任意单元格，可指定宽高（只给一个则保持比例）；
  尺寸从文件头解析，**不需要 Pillow**
- **健壮性**：XML 转义、非法控制字符清理、工作表名校验、ZIP 条目固定时间戳（输出可复现）

## API 速查

```python
book = Workbook()                      # 也可 Workbook("报表")
sheet = book.add_sheet("数据")          # 不给名字则自动 Sheet1 / Sheet2 …

sheet.write(row, column, value, style=None)     # 行列从 1 开始，列可写 "C" 或 3
sheet.append([...], style=None)                 # 写下一行，返回行号
sheet.merge("A1:C1", "标题", style=None)         # 可同时写入左上角
sheet.set_column_width("A", 24)                 # 或 set_column_width(1, 24)
sheet.add_image(png数据_或_路径, anchor="B2", width=None, height=None)

book.save("out.xlsx")                  # 也支持 save(buffer) / to_bytes()
```

## 文件是怎么拼出来的

xlsx 本质是一个装着 XML 的 ZIP。`zeroxlsx` 只写 Excel 真正需要的部件：

```
[Content_Types].xml · _rels/.rels · docProps/{core,app}.xml
xl/workbook.xml (+ rels) · xl/styles.xml · xl/worksheets/sheetN.xml
xl/drawings/drawingN.xml (+ rels)   ← 仅在有图片时
xl/media/imageN.png|jpeg
```

字符串用 `inlineStr` 写在单元格里（省掉共享字符串表）；图片用 `oneCellAnchor`
加显式 EMU 尺寸（96 dpi 下 1 px = 9525 EMU）。

## 测试

库本身零依赖，但测试**故意**用 openpyxl + Pillow 回读，逐项断言值、样式、合并、
列宽、公式、图片锚点 —— 以此证明产物能被主流实现正确读出。

```bash
pip install -e ".[dev]"
pytest
```

## 已知局限

只写不读；没有图表、数据透视、单元格内图片（浮动图片用 `add_image`）、条件格式、
数据验证、筛选、冻结窗格；也不是流式写出（整表在内存里构建）。

路线图：冻结窗格、自动筛选、流式写入器、单元格内图片，以及独立的 `zeroxlsx.read`。

## 许可

MIT，见 [LICENSE](LICENSE)。
