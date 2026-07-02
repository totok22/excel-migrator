# excel-migrator

把已填好内容的旧版 Excel 迁移到新版模板，保留格式、公式与图片。通过上下文匹配自动完成迁移，无法确定的单元格列入报告供人工确认。

> 由 BITFSAE 车队开发，最初用于 FSEC ESF（电气安全表）的版本迁移。



## 项目结构

```
src/excel_migrator/
├── core.py          # 单元格识别、SheetCache、两遍匹配引擎
├── images.py        # 图片复制、缩放、锚点处理
├── strict.py        # XML 层写回，处理 Excel 兼容性
├── pipeline.py      # 流程编排（加载 → 迁移 → 写出 → 报告）
├── profile_esf.py   # FSEC ESF 专用规则
├── server.py        # 本地 Web 服务器（stdlib http.server + SSE）
├── report.py        # 报告生成
└── cli.py           # 命令行入口
web/                 # 纯 HTML/CSS/JS 前端，无框架
templates/           # 内置空白模板
```



## 快速开始

需要 Python 3.10+。首次运行会自动创建虚拟环境并安装 `openpyxl` 和 `Pillow`。

| 系统 | 启动方式 |
|------|---------|
| macOS | 双击 `start.command`（首次右键 → 打开） |
| Windows | 双击 `start.bat` |
| Linux | `./start.sh` |

启动后浏览器自动打开 `http://127.0.0.1:8765/`。所有处理在本地完成，文件不会上传到任何服务器。



## 模式

- **通用模式（generic）**

  适用于结构相近、且用 `#FFFFCC99` 底色标记可填写区域的模板升版迁移。依赖 sheet 名、坐标、标签、表头、段落标题和文本相似度匹配。新旧文件需要是"模板升版"关系，即 sheet 和字段标签大体对应。

  未来要支持更多模板风格，可能的扩展方向：

  - 按用户指定的颜色或样式识别
  - 按非公式空白格识别（风险更高，容易误迁移）
  - 按模板中的占位符文本识别
  - 不限制输入格，扫描全部可写单元格（最宽松，误匹配风险最大）

  **目前没有做这些扩展的计划。**



- **FSEC ESF 2026 v2.2.4（esf）**

  在通用迁移基础上额外处理：
  - `接地 Grounding` sheet 按零件名称（`norm_text` 归一化后）做行匹配，而非坐标匹配
  - 碳纤维和外壳接地行按顺序对齐
  - `备用电池箱 Spare Accumulator` 图片锚点按 2026 模板的行偏移量重定位
  - 修复已知的模板公式断裂（`总览 Overview!K19`、`其他 Others!Z32`）
  - 使用内置空白模板 `templates/fsec_esf_template_2026_v2.2.4.xlsx`

## 高级设置

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `context_threshold` | 0.55 | 上下文 token Jaccard 阈值 |
| `fuzzy_threshold` | 0.70 | 标签/标题字符级相似度阈值 |
| `image_margin` | 0.92 | 图片缩放时保留的边距比例 |
| `filter_status` | true | 过滤公式生成的 Warning/OK 等状态文本 |
| `keep_instructional` | true | 保留模板中的图片放置提示等指导文字 |



## 输出

- **输出 Excel**：以新模板为骨架，填入旧版中可安全确认的内容。格式、公式、合并区、数据校验保持模板原样。
- **迁移报告**：多 sheet 的 Excel 报告，包含总览 KPI、需人工确认列表、匹配方法统计、单元格明细。



## 常见问题

- **有些单元格没填上？** 查看报告中的「需人工确认」sheet，列出了所有未自动填充的单元格及原因（未找到安全匹配 / 存在多处可能值）。

- **图片位置偏移？** 通用模式保持源文件原始锚点。新模板行高/列宽变化时图片会偏移，需在 profile 的 image targeter 中处理。

- **支持 .xls 吗？** 不支持，仅支持 `.xlsx`。请先用 Excel 另存为 `.xlsx`。

- **Windows 找不到 Python？** `start.bat` 依次尝试 `py -3`、`python`、`python3`。从 [python.org](https://www.python.org/downloads/) 安装时勾选 **Add python.exe to PATH**。



## 命令行

```bash
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# 通用模式
python -m excel_migrator.cli 旧版.xlsx 新模板.xlsx -o 输出.xlsx

# FSEC ESF 模式（使用内置模板）
python -m excel_migrator.cli 旧版.xlsx templates/fsec_esf_template_2026_v2.2.4.xlsx --profile esf
```

```
positional:
  source                旧版 Excel
  template              新版模板

options:
  -o, --output          输出路径（默认：旧文件名_migrated.xlsx）
  -r, --report          Excel 报告路径
  --profile {generic,esf}
  --overwrite           覆盖模板中已有的非占位内容
  --no-images           跳过图片迁移
  --keep-template-images  保留新模板中原有的图片
```



## 匹配算法

**前提**：模板必须用橙色底色（`#FFFFCC99`）标记用户填写区域。工具只迁移这些输入单元格中的值，公式单元格和状态文本（Warning/OK/Incomplete）不会被复制。不满足这个约定的 Excel 文件无法使用本工具。

迁移分两遍进行：

**第一遍：精确坐标匹配**

对目标模板中每个输入单元格，检查源文件同坐标是否有可迁移的值，同时验证上下文一致性（左侧标签、表头、段落标题的 Jaccard 相似度 ≥ 0.55）。通过则直接写入。

**第二遍：索引 + 模糊匹配**

对第一遍未填充的单元格，通过 `SheetCache.field_keys()` 生成多级索引键，按优先级依次查找：

1. `left_section` — 左侧标签 + 段落标题
2. `rowcol` — 左侧标签 + 段落 + 表头
3. `table` — 段落 + 表头 + 行偏移
4. `left` — 仅左侧标签
5. `fuzzy-context` — 字符级 Jaccard 相似度回退（阈值 0.70）

每个源单元格最多使用一次，避免一个值被写入多个位置。

**`SheetCache`**

`SheetCache` 在迁移开始前对每个 worksheet 预计算并缓存：合并单元格的 top-left 映射、静态文本（非输入、非公式的单元格内容）、左侧标签、最近表头、段落标题、上下文 token 集合。所有查找均为 O(1) 或 O(n) 一次性扫描，避免在迁移循环中重复遍历 worksheet。



## 输出写回机制

openpyxl 保存 xlsx 时会丢失部分 Excel 特有的结构（空样式单元格、namespace 声明、drawing namespace 前缀等），直接保存会导致 Excel 打开时弹出修复提示。

`strict.py` 通过直接操作 xlsx ZIP 包内的 XML 解决这个问题：

1. `build_strict_output`：以模板的 worksheet XML 为骨架，只把 staging workbook 中的单元格值 patch 进去，保留模板的格式、公式、合并区、数据校验。
2. `restore_empty_cells_from_reference`：把模板中存在但 openpyxl 写出时丢失的空样式单元格补回来。
3. `_ensure_ignorable_ns_declared`：检查 `mc:Ignorable` 引用的 namespace 前缀是否都有对应的 `xmlns:` 声明，缺失则注入，防止 Excel 报 corruption。
4. `_dedup_styles`：WPS 生成的模板常含重复 `cellXfs` 条目，Excel 打开时会触发修复提示，这里在写出前去重。



## 依赖

- [openpyxl](https://openpyxl.readthedocs.io/) — Excel 读写
- [Pillow](https://pillow.readthedocs.io/) — 图片处理

## License

[MIT](LICENSE) © 2026 totok22 / BITFSAE
