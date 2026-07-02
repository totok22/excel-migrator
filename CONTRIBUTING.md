# Contributing

目标：把旧模板中人工填写的内容迁移到新模板骨架。通用迁移不绑定任何具体赛事、车队或模板版本；模板特有的处理逻辑放到 profile。

## 代码结构

| 文件 | 作用 |
|------|------|
| `core.py` | 单元格识别、`SheetCache` 预计算缓存、两遍匹配引擎 |
| `images.py` | 图片复制、缩放（`fit_image_to_cell`）、锚点处理 |
| `strict.py` | XML 层写回，处理 openpyxl 与 Excel 的兼容性差异 |
| `pipeline.py` | 流程编排：加载 → profile → 迁移 → 图片 → 写出 → 报告 |
| `profile_esf.py` | FSEC ESF 2026 v2.2.4 专用规则 |
| `server.py` | stdlib `http.server` + SSE 进度推送，无第三方 web 框架 |
| `report.py` | Excel 报告生成 |

## 通用模式边界

通用模式只能依赖模板本身的信息：sheet 名、坐标、左侧标签、表头、段落标题、合并区布局、文本相似度、用户填写值、超链接、批注、图片锚点。

不要加入：特定 sheet 名的业务分支、特定坐标的修复、特定模板版本的图片重定位、某赛事或车队的字段含义判断。这些放到 profile。

## 新增 Profile

以新增 `profile_foo.py` 为例：

1. 新建 `src/excel_migrator/profile_foo.py`，实现专用迁移函数，例如 `migrate_foo(source_wb, target_wb, overwrite) -> list[CellAction]`。
2. 如需图片重定位，提供 `foo_image_targeter(source_ws, target_ws, img) -> str | anchor | None`：
   - 返回坐标字符串：图片缩放后放入该坐标所在的合并区域（`fit_image_to_cell`）
   - 返回 anchor 对象：保留原大小和偏移，只调整行列
   - 返回 `None`：回退到通用的 same-anchor 复制
3. 如需跳过通用迁移的 sheet，定义 `FOO_SKIP_SHEETS: set[str]`。
4. 在 `pipeline.py` 中根据 `opts.profile` 调用对应 profile。
5. 在 `cli.py` 和 `server.py` 允许新的 profile 名称。
6. 如有内置模板，放入 `templates/`，在 `server.py` 中显式声明路径。
7. 如需 Web 界面支持，在 `web/index.html` 和 `web/app.js` 增加选项。

profile 只处理确定的模板差异。无法确定的迁移留给通用报告提示人工确认。

## strict.py 注意事项

`build_strict_output` 直接操作 xlsx ZIP 包内的 XML，不经过 openpyxl 的高层 API。修改时需注意：

- **namespace 声明**：`_ensure_ignorable_ns_declared` 会检查 `mc:Ignorable` 引用的前缀是否都有 `xmlns:` 声明。openpyxl 的 `ET.tostring` 只输出实际用到的前缀，但 `mc:Ignorable` 可能引用已被移除元素的前缀，缺失会导致 Excel 报 corruption。
- **array formula**：`_patch_cell` 跳过 `t="array"` 的公式单元格，因为 openpyxl 会丢失缓存的 `<v>` 值，保留模板原始结构更安全。
- **drawing namespace**：`_fix_drawing_xml` 修复 openpyxl 的已知 bug：`<avLst />` 应为 `<a:avLst />`，否则 Excel 报 drawing corrupt。
- **样式去重**：`_dedup_styles` 处理 WPS 生成模板中常见的重复 `cellXfs` 条目，避免 Excel 打开时触发修复提示。

## 图片迁移

多个图片位于同一图片区时，优先返回调整后的 anchor（保留相对位置），而非全部返回坐标字符串（会导致所有图片居中叠在一起）。

参考 `profile_esf.py` 中 `_spare_accumulator_image_anchor` 的实现：按源图片的行列范围判断所属区块，用 `_shift_image_anchor_rows` 整体平移 anchor，保留原始的列偏移和图片尺寸。

## 验证

提交前至少运行：

```bash
python3 -m compileall -q src
node --check web/app.js
git diff --check
```

涉及 xlsx 输出时：

```bash
unzip -t output.xlsx
python3 -c "from openpyxl import load_workbook; wb = load_workbook('output.xlsx'); print(wb.sheetnames); wb.close()"
```

修复 Excel 兼容性或版式问题时，需用 Microsoft Excel 实际打开输出文件确认不弹修复提示。

## 提交约定

- 不要提交迁移结果、报告文件或 `~$` Office 锁文件。
- 内置模板必须是空白模板，不能包含车队私有填写数据。
- README 写用户需要知道的行为；实现细节写在本文件。
