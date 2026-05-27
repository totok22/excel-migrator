# Excel Migrator

把已填好内容的旧版 Excel **自动迁移**到新版模板，保留格式、公式与图片。

适用场景：模板升版后，需要把旧版中手动填写的数据搬到新模板里，但手动复制粘贴容易遗漏或破坏格式。

> 由 BITFSAE 车队开发，最初用于 FSC ESF（电气安全表）的版本迁移。

---

## 快速开始

### 前置条件

| 条件 | 说明 |
|------|------|
| **Python 3.10+** | [下载地址](https://www.python.org/downloads/)。macOS 可用 `brew install python`，Windows 推荐从官网安装并勾选 "Add to PATH" |
| **网络连接** | 仅首次启动时需要（自动安装 `openpyxl` 和 `Pillow` 两个依赖） |

不需要安装 Excel，不需要管理员权限。

### 一键启动

把整个 `excel-migrator/` 文件夹放到本地任意位置，然后双击对应系统的启动脚本：

| 系统 | 启动方式 |
|------|---------|
| **macOS** | 双击 `start.command`（首次可能需要右键 → 打开） |
| **Windows** | 双击 `start.bat` |
| **Linux** | 终端运行 `./start.sh` |

首次启动会自动创建虚拟环境并安装依赖（约 30 秒），之后每次启动秒开。

启动成功后浏览器会自动打开 `http://127.0.0.1:8765/`，看到 Web 界面即可使用。

### 使用步骤

1. **选择模式** — 通用模式适用于任意 Excel 模板迁移；FSC ESF 模式额外启用接地表等专用规则
2. **上传文件** — 拖入或选择「旧版 Excel」和「新版模板」
3. **点击开始** — 等待进度条完成
4. **下载结果** — 下载迁移后的 Excel 和迁移报告

所有处理在本地完成，文件不会上传到任何服务器。

---

## 命令行使用

如果不想用 Web 界面，也可以直接用命令行：

```bash
# 激活虚拟环境
source .venv/bin/activate  # macOS/Linux
# 或 .venv\Scripts\activate  # Windows

# 运行迁移
python -m excel_migrator.cli 旧版.xlsx 新模板.xlsx -o 输出.xlsx --report 报告.xlsx

# FSC ESF 模式
python -m excel_migrator.cli 旧版.xlsx 新模板.xlsx --profile esf
```

完整参数：

```
positional arguments:
  source                旧版 Excel 文件路径
  template              新版模板 Excel 路径

options:
  -o, --output          输出 xlsx 路径（默认：旧文件名_migrated.xlsx）
  -r, --report          Excel 报告路径（默认：输出文件名_报告.xlsx）
  --profile {generic,esf}  迁移配置（默认 generic）
  --overwrite           覆盖新版模板中已有的非占位内容
  --no-images           不迁移图片
  --keep-template-images  保留新模板中原有的图片
```

---

## 输出说明

| 文件 | 内容 |
|------|------|
| **输出 Excel** | 以新模板为骨架，填入旧版中可安全确认的内容。格式、公式、合并区、数据校验保持模板原样 |
| **迁移报告** | 多 sheet 的 Excel 报告：总览 KPI、需人工确认列表、按 sheet 分布、匹配方法统计、单元格明细 |

---

## 项目结构

```
excel-migrator/
├── README.md
├── LICENSE                  # MIT
├── requirements.txt
├── start.command            # macOS 启动脚本
├── start.sh                 # Linux 启动脚本
├── start.bat                # Windows 启动脚本
├── web/                     # 前端（纯 HTML/CSS/JS，无框架）
│   ├── index.html
│   ├── style.css
│   └── app.js
└── src/excel_migrator/      # Python 后端
    ├── __init__.py
    ├── cli.py               # 命令行入口
    ├── server.py            # 本地 Web 服务器
    ├── pipeline.py          # 主流程编排
    ├── core.py              # 单元格匹配引擎
    ├── images.py            # 图片迁移
    ├── strict.py            # 模板严格写回
    ├── profile_esf.py       # FSC ESF 专用规则
    └── report.py            # 报告生成
```

---

## 匹配策略

迁移分两遍进行：

### 第一遍：精确匹配

对每个目标模板中的输入单元格（橙色底色），检查源文件同坐标位置是否有可迁移的值，并验证上下文（左侧标签、表头、段落标题）是否一致。一致则直接写入。

### 第二遍：智能匹配

对第一遍未填充的单元格，通过多级索引查找：

1. **left_section** — 左侧标签 + 段落标题
2. **rowcol** — 左侧标签 + 段落 + 表头联合
3. **table** — 段落 + 表头 + 行偏移
4. **left** — 仅左侧标签
5. **fuzzy-context** — 模糊文本相似度回退

每个源单元格最多只会被使用一次，避免一个值被错误地写入多个位置。

### 安全机制

- 公式不会被复制（只迁移用户手动填写的值）
- 公式生成的状态文本（Warning、OK、Incomplete 等）会被过滤
- 模板中的指导性文字（如图片放置提示）默认保留
- 未能安全匹配的单元格会列入报告的「需人工确认」列表

---

## 高级设置

Web 界面中点击「高级设置」可调整：

| 参数 | 默认值 | 说明 |
|------|--------|------|
| 上下文匹配阈值 | 0.55 | 上下文 token Jaccard 相似度阈值，越低越宽松 |
| 模糊文本匹配阈值 | 0.70 | 标签/标题的字符级相似度阈值 |
| 图片边距比例 | 0.92 | 图片缩放时保留的边距（仅对需要重定位的图片生效） |
| 状态值过滤 | 开 | 跳过公式生成的 Warning/OK 等状态文本 |
| 保留指导性文字 | 开 | 保留模板中的图片放置提示等指导文字 |

---

## 常见问题

**Q: 迁移后有些单元格没有填上？**

查看迁移报告中的「需人工确认」sheet，里面列出了所有未自动填充的单元格及原因。常见原因：新模板结构变化较大导致上下文不匹配，需要手动复制。

**Q: 图片位置不对？**

普通图片会保持源文件中的原始位置和大小。如果新模板的行高/列宽与旧版不同，图片可能看起来偏移。可以在 Excel 中手动调整。

**Q: 支持 .xls 格式吗？**

不支持。仅支持 `.xlsx`（Office 2007+）格式。如果有 `.xls` 文件，请先用 Excel 另存为 `.xlsx`。

**Q: 可以反复运行吗？**

可以。每次运行都是从原始模板重新生成，不会累积错误。

---

## 技术依赖

- Python 3.10+
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel 读写
- [Pillow](https://pillow.readthedocs.io/) — 图片处理

---

## License

[MIT](LICENSE) © 2026 totok22 / BITFSAE
