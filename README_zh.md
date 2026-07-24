# excel2json

将 Excel (`.xls`/`.xlsx`) 文件转换为 **JSON** 和 **CSV**。

Go 语言重写的 [excel2json](https://github.com/neil3d/excel2json) C# 工具，并额外支持 CSV 导出。

---

## 安装

```bash
go install github.com/actfuns/excel2json@latest
```

或从源码编译：

```bash
git clone https://github.com/actfuns/excel2json.git
cd excel2json
go build -o excel2json .
```

---

## 使用

```bash
excel2json [options] <输入文件或目录...> [-o <路径>] [-f json|csv]
```

输入可以是文件路径或目录（目录会自动收集所有 `.xls`/`.xlsx` 文件）。

### 示例

```bash
# JSON 输出到标准输出
excel2json data.xlsx

# CSV 输出到标准输出
excel2json data.xlsx -f csv

# 美化 JSON
excel2json data.xlsx -p

# 写入目录（每个 sheet → 独立文件）
excel2json data.xlsx -o ./output/

# 合并所有 sheet 到一个文件
excel2json data.xlsx -o out.json -m

# 使用自定义列作为字典 key
excel2json data.xlsx -k cs_id

# 跳过前 2 行开始读数据
excel2json data.xlsx -s 2

# 排除指定前缀的 sheet 和列
excel2json data.xlsx -x S_ -x cs_

# 自定义输出文件名
excel2json data.xlsx --name file              # 只用文件名
excel2json data.xlsx --name sheet             # 只用 sheet 名
excel2json data.xlsx --name-tpl 'my_{file}_{sheet}'  # 自定义模板
```

### 选项

| 短 | 长 | 默认值 | 说明 |
|---|---|---|---|
| | `inputs...` | **必填** | Excel 文件或目录 |
| `-o` | `--out` | `""` | 输出路径（文件或目录） |
| `-f` | `--format` | `json` | 输出格式：`json` 或 `csv` |
| | | **数据解析** |
| `-n` | `--name-row` | `0` | 列名所在行号（从 0 开始） |
| `-t` | `--type-row` | `-1` | 类型注释所在行号（-1 禁用） |
| `-s` | `--skip-rows` | `1` | 数据前跳过的行数 |
| `-e` | `--encoding` | `utf8-nobom` | 输出文件编码 |
| `-d` | `--date-format` | `yyyy/MM/dd` | 日期格式字符串 |
| | | **字段处理** |
| `-l` | `--lowcase` | `false` | 字段名转小写 |
| `-x` | `--exclude` | `nil` | 跳过指定前缀的 sheet 和列（可重复） |
| | `--cell-json` | `false` | 解析单元格中的 JSON 字符串 |
| | `--all-string` | `false` | 全部输出为字符串 |
| | | **输出格式** |
| `-a` | `--array` | `false` | 导出为数组（默认以第一列值为 key） |
| `-m` | `--merge` | `false` | 合并所有 sheet 到一个文件（默认拆分） |
| | `--name-tpl` | `{file}_{sheet}` | 输出文件名模板，支持 `{file}` `{sheet}` |
| | `--name` | `""` | 文件名预设：`both`, `file`, `sheet` |
| `-p` | `--pretty` | `false` | JSON 美化输出（Tab 缩进） |
| `-k` | `--key` | `""` | 用作字典 key 的列名（默认第一列） |

---

## 输出格式

### JSON — 字典模式（默认）

每行以第一列值为 key：

```json
{
  "1": { "name": "张三", "age": 30 },
  "2": { "name": "李四", "age": 25 }
}
```

### JSON — 数组模式（`-a`）

每行为数组元素：

```json
[
  { "id": 1, "name": "张三", "age": 30 },
  { "id": 2, "name": "李四", "age": 25 }
]
```

### 多个 Sheet

默认拆分为独立文件（文件名 = `{文件名}_{sheet名}.json`）。

使用 `-m` 可合并到一个文件：

```json
{
  "员工": { "1": { "name": "张三" } },
  "部门": { "1": { "name": "技术部" } }
}
```

### CSV

- **单 sheet**：写入指定路径
- **多 sheet**：自动创建 `{文件名}_{sheet名}.csv` 文件

---

## 路由

| 输入 | 无 `-o` | `-o <文件>` | `-o <目录>/` |
|---|---|---|---|
| 单个文件 | stdout | 写入文件 | `{目录}/{文件名}.{格式}` |
| 多个文件 | stdout（依次） | 当作目录 | `{目录}/{文件名}.{格式}` |
| 目录 | stdout（依次） | 当作目录 | 同上 |

---

## 特性

- **并行处理**：文件并发处理（上限 `2 × CPU 核心数`）
- **前缀排除**：使用 `-x`（前缀匹配）跳过 sheet 和列；可重复使用
- **智能单元格解析**：单元格中的 JSON 字符串自动反序列化（`--cell-json`）
- **全字符串模式**：关闭所有解析，输出原始字符串（`--all-string`）
- **字段名小写**：`-l` 转小写
- **数值去零**：`88.0` → `88`，`3.14` → `3.14`
- **浮点精度修正**：`15699.999999999998` 自动归整为 `15700`
- **空值默认值**：以同列第一个非空值推断类型（数值→`0`，字符串→`""`）
- **编码**：默认 UTF-8 无 BOM
- **.xls 支持**：通过 [grate](https://github.com/pbnjay/grate) 读取旧版 BIFF / WPS 文件

---

## 依赖

- [excelize](https://github.com/xuri/excelize) — XLSX 读取
- [grate](https://github.com/pbnjay/grate) — XLS (BIFF) 读取
- [cobra](https://github.com/spf13/cobra) — CLI 框架
- [errgroup](https://pkg.go.dev/golang.org/x/sync/errgroup) — 并发控制

---

## 许可证

MIT
