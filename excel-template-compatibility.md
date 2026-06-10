# Excel 多模板兼容输入处理

处理多种不同格式的 Excel 文件输入时的通用方法论和实现指南。

## 1. 支持的文件格式与读取库选择

| 格式 | 扩展名 | 推荐库 | 说明 |
|------|--------|--------|------|
| 旧版 Excel | `.xls` | `xlrd`（需 `formatting_info=True`） | BIFF 格式，API 与 openpyxl 完全不同 |
| 现代 Excel | `.xlsx` | `openpyxl` | OOXML 格式 |
| 宏启用 Excel | `.xlsm` | `openpyxl` | 与 xlsx 处理方式相同 |

### 格式路由模式

根据文件扩展名路由到不同的读取策略，是兼容多格式的基本模式：

```python
def read_excel(file_path: Path) -> list[SheetData]:
    if file_path.suffix.lower() == '.xls':
        return _read_xls(file_path)     # xlrd 路径
    # .xlsx / .xlsm → openpyxl 路径
    return _read_xlsx(file_path)
```

## 2. 统一数据模型

兼容层的核心原则：**不同格式的读取逻辑各自独立，但输出归一化为统一的数据结构**，下游模块完全不感知输入格式差异。

```python
@dataclass
class CellData:
    value: str | None
    is_strikethrough: bool

@dataclass
class SheetData:
    name: str
    rows: list[list[CellData]]
```

所有格式的读取函数都必须返回 `list[SheetData]`，这样清洗、分析、输出等后续环节只需面对一种数据结构。

## 3. .xlsx/.xlsm 读取要点（openpyxl）

### 3.1 双重加载模式

openpyxl 的 `data_only` 和 `rich_text` 模式互斥，需要同时获取计算值和格式信息时要加载两次：

```python
wb_data = load_workbook(file_path, data_only=True)   # 公式单元格返回计算结果
wb_rich = load_workbook(file_path, rich_text=True)    # 单元格值为 CellRichText 对象
```

### 3.2 富文本部分删除线处理

一个单元格内可能只有部分文本被标记为删除线，需要逐段检查 `CellRichText`：

- `TextBlock.font.strike is True` → 该段有删除线
- `TextBlock.font.strike is False` → 该段无删除线
- `TextBlock.font.strike is None`（font 对象存在）→ 有独立格式覆盖，视为无删除线
- `TextBlock.font is None` → 无格式信息，继承单元格级别的 `font.strike`
- 纯字符串片段（`isinstance(part, str)`）→ 继承单元格级别

### 3.3 对角叉检测

对角叉（X 形交叉线）可作为删除标记：

```python
has_diagonal_cross = (
    cell.border
    and cell.border.diagonalUp
    and cell.border.diagonalDown
)
```

## 4. .xls 读取要点（xlrd）

### 4.1 基本读取

```python
wb = xlrd.open_workbook(str(file_path), formatting_info=True)
```

`formatting_info=True` 是关键参数，启用后才能访问字体、边框等格式信息。

### 4.2 值类型转换

xlrd 的单元格类型需要手动转换为字符串：

```python
if cell.ctype == xlrd.XL_CELL_EMPTY:
    # 空单元格
elif cell.ctype == xlrd.XL_CELL_NUMBER:
    value = str(int(raw)) if raw == int(raw) else str(raw)  # 整数不带小数点
else:
    value = str(raw) if raw != '' else None
```

### 4.3 删除线检测

xlrd 通过 XF（扩展格式）记录链访问字体信息：

```python
xf_idx = ws.cell_xf_index(ri, ci)
xf = wb.xf_list[xf_idx]
cell_font = wb.font_list[xf.font_index]
cell_strike = bool(cell_font.struck_out)
```

### 4.4 部分删除线（Rich Text Runs）

xlrd 使用 `rich_text_runlist_map` 存储富文本信息：

```python
runs = ws.rich_text_runlist_map.get((ri, ci))
# runs: [(start_char, font_idx), ...]
```

每个 run 表示从 `start_char` 开始使用 `font_idx` 指定的字体。处理步骤：
1. 计算每个 run 的文本范围
2. 检查对应字体的 `struck_out` 属性
3. 只保留非删除线部分
4. 第一个 run 之前的文本继承单元格级别的字体属性

### 4.5 对角叉检测

```python
has_diagonal = (
    xf.border.diag_line_type != 0
    and xf.border.diag_colour_index != 0x7FFF  # 非未设定色
)
```

## 5. 文件扫描与过滤

扫描输入目录时的通用实践：

```python
patterns = ['*.xlsx', '*.xls', '*.xlsm']
files = []
for pattern in patterns:
    files.extend(Path(input_dir).glob(pattern))
# 排除 Excel 临时文件（~$ 开头）
files = [f for f in files if not f.name.startswith('~$')]
return sorted(files, key=lambda f: f.name)
```

## 6. 数据清洗统一层

对所有格式的数据执行统一清洗，消除格式差异后的残留噪声：

- 删除线单元格 → 值置为空字符串
- 非删除线单元格 → 保留原始文本，去除首尾空白
- 全空行 → 移除
- 清洗后无数据的 sheet → 跳过

## 7. 用 AI 动态适配不同模板结构

当不同模板的表格结构（列布局、数据起始行、sheet 命名）各不相同时，可以用 AI 二阶段分析自动适配，避免为每种模板硬编码列映射：

### Phase 1：结构识别

- 输入：所有 sheet 的前 N 行，格式化为 `[列号]值` 便于 AI 理解表格结构
- AI 识别：文档元信息、数据 sheet 名、数据起始行、关键列位置
- 通过 Tool Call 返回结构化结果

### Phase 2：数据提取

- 根据 Phase 1 识别的列位置过滤数据列，减少 token 消耗
- 按固定行数分块发送，避免超出上下文窗口
- AI 从每个块中提取结构化记录

## 8. 容错设计原则

- **单文件失败隔离**：一个文件处理失败不中断整体批量流程
- **空 sheet 跳过**：清洗后无数据的 sheet 自动跳过，记录日志
- **缺失字段填充**：解析结果中缺失的字段用空字符串填充，不中断流程
- **AI 调用重试**：指数退避重试，应对临时性网络或服务故障
- **格式信息异常兜底**：格式信息读取失败时默认为无特殊格式，避免因个别单元格异常导致整个文件失败

## 9. 扩展新模板格式的步骤

1. **新的文件扩展名**：在扫描模块的 `patterns` 列表中添加
2. **新的读取库**：在读取模块中添加新的路由分支，确保输出为统一数据模型
3. **新的表格布局**：如果使用 AI 适配，无需修改代码；如果硬编码，需添加新的列映射配置
4. **新的清洗规则**：在清洗模块中扩展
5. **新的输出字段**：在数据模型中添加字段，同步更新解析和输出模块
