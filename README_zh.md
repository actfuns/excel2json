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
excel2json [options] <输入文件或目录...> [-o <路径>] [--format json|csv]
```

输入可以是文件路径或目录（目录会自动收集所有 `.xls`/`.xlsx` 文件）。

### 示例

```bash
# JSON 输出到标准输出
excel2json data.xlsx

# CSV 输出到标准输出
excel2json data.xlsx --format csv

# 美化 JSON 输出到标准输出
excel2json data.xlsx --pretty

# 写入文件
excel2json data.xlsx -o out.json

# 写入目录
excel2json data.xlsx -o ./output/ --pretty

# 多文件写入目录
excel2json a.xlsx b.xlsx -o ./output/

# 输入为目录（自动扫描所有 .xls/.xlsx）
excel2json ./data/ -o ./output/
```

### 选项

| 参数 | 默认值 | 说明 |
|---|---|---|
| `inputs...` | **必填** | Excel 文件或目录 |
| `-o`, `--out` | `""` | 输出路径（单文件时写文件，多文件时当目录） |
| `--format` | `json` | 输出格式：`json` 或 `csv` |
| `--header` | `1` | 表头行数 |
| `-c`, `--encoding` | `utf8-nobom` | 输出文件编码 |
| `-l`, `--lowcase` | `false` | 字段名转小写 |
| `-a`, `--array` | `false` | 导出为数组（默认以第一列值为 key 的字典） |
| `-d`, `--date` | `yyyy/MM/dd` | 日期格式 |
| `-s`, `--sheet` | `false` | 即使只有一个 sheet 也强制按 sheet 名包装 |
| `-x`, `--exclude_prefix` | `""` | 跳过指定前缀的 sheet/列 |
| `--cell_json` | `false` | 解析单元格中的 JSON 字符串 |
| `--all_string` | `false` | 全部输出为字符串（关闭 JSON 和数值解析） |
| `--pretty` | `false` | JSON 美化输出（Tab 缩进） |

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

自动按 sheet 名包装：

```json
{
  "员工": { "1": { "name": "张三", "age": 30 } },
  "部门": { "1": { "name": "技术部" } }
}
```

使用 `-s` 即使只有一个 sheet 也强制包装。

### CSV

- **单 sheet**：写入指定路径
- **多 sheet / 目录输出**：创建 `{sheet名}.csv` / `{文件名}_{sheet名}.csv` 文件

---

## 输出路由

| 输入 | 无 `-o` | `-o <文件>` | `-o <目录>/` |
|---|---|---|---|
| 单个文件 | stdout | 写入文件 | `{目录}/{文件名}.{格式}` |
| 多个文件 | stdout（依次） | 当作目录 | `{目录}/{文件名}.{格式}` |
| 目录 | stdout（依次） | 当作目录 | 同上 |

---

## 特性

- **并发处理**：文件并行处理（上限 `2 × CPU 核心数`）
- **智能单元格解析**：单元格中的 JSON 字符串自动反序列化（`--cell_json`）
- **全字符串模式**：关闭所有解析，输出原始字符串（`--all_string`）
- **字段名小写**：`--lowcase` 转为小写字母
- **排除前缀**：跳过指定前缀的 sheet 或列名（`-x`）
- **数值去零**：`88.0` → `88`，`3.14` → `3.14`
- **空值默认值**：以同列第一个非空值推断类型（数值→`0`，字符串→`""`）
- **编码**：默认 UTF-8 无 BOM

---

## 依赖

- [excelize](https://github.com/xuri/excelize) — Excel 文件读取
- [cobra](https://github.com/spf13/cobra) — CLI 框架
- [errgroup](https://pkg.go.dev/golang.org/x/sync/errgroup) — 并发控制

---

## 许可证

MIT