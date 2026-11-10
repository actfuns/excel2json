# excel2json

将 Excel (`.xls`/`.xlsx`) 文件转换为 **JSON** 和 **CSV**。

Go 语言重写的 [excel2json](https://github.com/neil3d/excel2json) C# 工具，并额外支持 CSV 导出。

---

## 安装

```bash
go install github.com/actfuns/excel2json/cmd/excel2json@latest
```

或从源码编译：

```bash
git clone https://github.com/actfuns/excel2json.git
cd excel2json
go build -o excel2json ./cmd/excel2json   # or: make build
```

---

## 使用

```bash
excel2json [options] <输入文件或目录...> [-o <路径>] [-f json|csv]
excel2json --version
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
excel2json data.xlsx -o out.json

# 使用自定义列作为字典 key
excel2json data.xlsx -k cs_id

# 指定数据起始行（索引 2；默认 -1 = 自动）
excel2json data.xlsx --data-row 2

# 排除指定前缀的 sheet 和列
excel2json data.xlsx -x S_ -x cs_

# 自定义输出文件名
excel2json data.xlsx --name-tpl 'my_{file}_{sheet}'  # 自定义模板

# 强类型列（类型行 = 第 2 行，数据从第 3 行开始）
excel2json data.xlsx -t 1 --required-by-default
```

### 选项

| 短 | 长 | 默认值 | 说明 |
|---|---|---|---|
| | `inputs...` | **必填** | Excel 文件或目录 |
| `-o` | `--out` | `""` | 输出路径（文件或目录） |
| `-f` | `--format` | `json` | 输出格式：`json` 或 `csv` |
| | | **数据解析** |
| `-n` | `--header-row` | `0` | 列名所在行号（从 0 开始） |
| `-t` | `--type-row` | `-1` | 类型注释所在行号（-1 禁用） |
| `-r` | `--required-by-default` | `false` | 表头未加 `!`/`?` 标记的字段默认必填（空单元格报错） |
| `-s` | `--data-row` | `-1` | 数据起始行号（0 开始，-1 = 自动） |
| `-e` | `--encoding` | `utf8-nobom` | 文件输出编码：`utf8`、`utf8-nobom`、`utf8-bom`、`gbk`、`gb18030` |
| `-d` | `--date-format` | `yyyy/MM/dd` | 日期格式字符串 |
| | | **字段处理** |
| `-l` | `--lowercase` | `false` | 字段名转小写 |
| `-x` | `--exclude` | `nil` | 跳过指定前缀的 sheet 和列（可重复） |
| | `--tag` | `""` | 只导出带该标签的 sheet/列；无标签视为公共 |
| | `--strict-tags` | `false` | 配合 `--tag`，导出无标签的 sheet/列时报错 |
| | `--strict-fields` | `false` | 输出字段名重复时报错（默认后者覆盖前者） |
| | `--parse-cell-json` | `false` | 解析无类型列单元格中的 JSON 字符串 |
| | `--stringify-all` | `false` | 无类型列全部输出为字符串 |
| | `--types` | `""` | types.yaml / types.yml 路径（省略时自动查找） |
| | | **输出格式** |
| | `--list` | `false` | JSON 导出为数组（默认以第一列值为 key） |
| `-m` | `--merge-sheets` | `false` | 合并所有 sheet 到一个文件（JSON 对象或 CSV 分段） |
| | `--name-tpl` | `{file}_{sheet}` | 拆分输出文件名模板，支持 `{file}` `{sheet}`（JSON 和 CSV） |
| `-p` | `--pretty` | `false` | JSON 美化输出（Tab 缩进） |
| `-k` | `--key` | `""` | 用作字典 key 的列名（默认第一个导出的列） |
| | `--unique-keys` | `false` | 字典 key 重复或为空时报错（默认后者覆盖前者） |

行号参数（`--header-row`、`--type-row`、`--data-row`）按 Excel 实际行号计算（0 开始）。中间空行会保留用于定位表头/类型行，数据区空行会跳过。`--data-row=-1` 表示取表头/类型行之后的第一行。

---

### 标签（客户端/服务端字段）

标签用于区分导出目标，同时保持 JSON key 不变：

```text
reward@c        # 仅客户端
reward@s        # 仅服务端
reward@c,s      # 双端
reward          # 无标签 = 公共
Shop@c          # sheet 级标签
```

```bash
excel2json --tag c -x '#' data.xlsx -o dist/client/
excel2json --tag s -x '#' data.xlsx -o dist/server/
```

- `@c,s` 只作为元数据，JSON key 仍是 `reward`；sheet 文件名去掉后缀（`Shop@c` → `Shop.json`）
- `-x/--exclude` 先执行，是面向所有端的硬排除（例如用 `-x '#'` 排除注释）
- 无标签的 sheet/列视为公共，任何 `--tag` 都会导出
- `--strict-tags` 下，被导出的 sheet/列若没有标签则报错
- 字典模式默认使用第一个导出的列作为 key
- 同一 target 下两个列最终 key 相同会报错
- `ref(...)` 始终基于完整工作簿校验，不受 `--tag`/`-x` 导出过滤影响
- 没有匹配列的工作簿/Sheet 会跳过；全部无匹配时 `--tag` 会给出明确错误
- 标签允许字母、数字、`_`、`-`；字段名末尾的 `@...` 保留给标签使用
- `--strict-fields` 可让普通重复字段名也报错
- 标签可以和必填标记组合：`reward@c!`、`reward!@s`
- 不传 `--tag` 时导出全部 sheet/列；但去掉标签后同名 key 冲突仍会报错

---

## 输出格式

### JSON — 字典模式（默认）

每行以第一列值为 key。默认重复 key 后面覆盖前面；加 `--unique-keys` 才会报错。需要保留所有行时用 `--list` 导出数组：

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

- **`-o file.csv`**：所有 sheet 合并写入一个文件
- **`-o dir/`**：每个 sheet 按 `--name-tpl` 生成独立文件（默认 `{file}_{sheet}.csv`），只剩一个 sheet 时也一样

---

## 路由

| 输入 | 无 `-o` | `-o <文件>` | `-o <目录>/` |
|---|---|---|---|
| 单个文件 | stdout | 写入单个文件（多 sheet 自动包装/分段） | 按 `--name-tpl` 拆分，或 `-m` 合并为 `{目录}/{文件名}.{格式}` |
| 多个文件 | stdout（依次） | 当作目录 | 按 `--name-tpl` 拆分，或 `-m` 合并为 `{目录}/{文件名}.{格式}` |
| 目录 | stdout（依次） | 当作目录 | 同上 |

`-o <文件>` 表示「一个输出文件」，多个 sheet 时会合并。没有扩展名的路径（如 `-o out/conf`）或以 `/` 结尾的路径视为目录；目录默认拆分，`--merge-sheets` 才合并。

---

## 类型注解

通过 `-t/--type-row` 指定类型注解所在行（从 0 开始），该列会按类型严格转换和校验。

| 分类 | 类型 |
|---|---|
| 基础类型 | `bool`、`string`、`int`、`int8`、`int16`、`int32`、`int64`、`uint`、`uint8`、`uint16`、`uint32`、`uint64`、`uintptr`、`float32`、`float64`、`byte`、`rune`、`any`、`interface{}` |
| 数组 / 切片 | `[]T`、`[N]T` |
| Map | `map[K]V` |
| 额外别名 | `date`、`float`（= `float64`）、`object`（= `map[string]any`）、`list`（= `[]any`）、`null` |

复合类型要求单元格内容为合法 JSON，例如 `[1,2,3]` 或 `{"a":"b"}`。数组长度、元素类型、数值范围、map 的 key 和 value 都会校验。

`any` / `interface{}` 保持旧的自动解析行为：自动识别数字，只有加 `--parse-cell-json` 才解析 JSON。`date` 解析失败时会原样输出字符串。

### 必填 / 可空（表头）

空值只由表头标记或 `--required-by-default` 决定：

| 表头 | 含义 | 输出 key |
|---|---|---|
| `id!` | 必填，空单元格报错 | `id` |
| `name?` | 可空，允许空单元格 | `name` |

不加标记时使用 `-r/--required-by-default` 的全局默认值。空值输出 `0`、`false`、`""`，复合类型输出 `null`。

### when(...) 条件守卫

在规则开头加 `when(...)`，条件成立时才执行该规则：

```text
when(type=1) effects:[]int:ref(effect.id)
when(type!=0, level>=10) int:gt(0)
when(typeof($) = [3]int) ...
when(typeof(kind.a) = int) ...
```

- `type` — 其他列的原始值
- `kind.a` / `kind[0]` — 对其他列做 JSONPath
- `$.a` — 对当前单元格做 JSONPath
- `$` — 当前单元格
- `typeof(ref) = <类型表达式>` — 结构化 JSON 类型判断
- 运算符：`=`、`==`、`!=`、`>`、`>=`、`<`、`<=`；`typeof` 只支持 `=`、`==`、`!=`
- `when(...)` 内逗号 = AND；`;` 仍然分隔规则
- `and(...)`、`or(...)`、`not(...)` 可嵌套；`not(a,b)` = `NOT(a AND b)`
- `type(...)` 是 `typeof(...)` 的别名
- 类型表达式支持 `[3]int`、`[]T`、`map[K]V`、`[2][2]int` 等
- 原始单元格不是合法 JSON 时按 `string` 处理；`TRUE`/`FALSE` 识别为布尔值
- 列不存在或需要 JSON 但内容非法时返回错误；路径无匹配表示条件不成立；多匹配时任意一个满足即可

```text
when(type=1) effects:[]int:ref(effect.id)
when(typeof($) = int) gte(0)
when(typeof($) = string) len(2,4)
when(type=1, typeof(kind.a) = [3]int) ...
when(or(type=1, level>=10)) ...
when(not(type=1)) ...
```

### 校验规则

```ebnf
fieldSpec  = rule { ";" rule }
rule       = [ "when(" conditions ")" ] [ selector ":" ] [ type [ ":" checks ] ]
conditions = condition { "," condition }
checks     = check { "," check }
```

| 语法 | 含义 |
|---|---|
| `,` | 多条校验同时满足（AND） |
| `;` | 多条规则同时生效（AND，选择器允许重叠） |
| `when(...)` | 规则守卫，所有条件都满足才执行（AND） |
| 选择器省略 | 作用于字段本身 |
| 类型省略 | 按选择器产出的元素原值校验 |

选择器是 RFC 9535 JSONPath 的子集，支持 `$`、`.field`、`['field']`/`["field"]`、`[0]`、`[-1]`、`[*]`、`[start:stop:step]`、`[i,j,k]`、`..field`。多匹配按标准 nodelist 语义**扁平拼接**。

```text
int:gt(0),lt(100)
when(typeof($) = int) gt(0);when(typeof($) = string) len(1)
$[*][1::3]:string:ref(item.id)
$[0]:int:gte(0);$[2]:int32:gt(0)
[4]int32:pad,gte(0)
```

- 选择器只负责**定位**规则要校验的值，不是输出选择器；输出始终是整个字段。
- 所有规则都**原位写回**：选中的节点按规则转换后写回原结构，未选中的部分保持原样。
- 规则按书写顺序执行；重叠位置以后面的规则为准。
- 守卫不成立的规则跳过；如果所有规则都被跳过，字段保持原值。
- 选择器无匹配时按标准 JSONPath 返回空结果，不报错。

```text
输入 [[1,"a",10.0],[2,"b",20.0]]
规则 $[*][2::3]:int
输出 [[1,"a",10],[2,"b",20]]
```

数组上的标量校验器（`gt`、`len` 之外的字符串/数值校验、`ref` 等）自动作用于**每个元素**；`len`、`minlen`、`maxlen`、`unique`、`nonempty` 作用于整个集合。

### 校验器

| 分类 | 校验器 |
|---|---|
| 数值 | `gt(n)`、`gte(n)`、`lt(n)`、`lte(n)`、`between(min,max)`、`eq(v)`、`ne(v)`、`positive`、`negative`、`nonzero`、`multipleof(n)` |
| 字符串 | `len(n)`、`len(min,max)`、`minlen(n)`、`maxlen(n)`、`email`、`phone`、`phone(e164)`、`url`、`uuid`、`regex(pattern)`、`prefix(s)`、`suffix(s)`、`contains(s)`、`alpha`、`alnum`、`numeric` |
| 枚举 | `oneof(a,b,c)` / `in(...)`、`notoneof(...)` / `notin(...)` |
| 集合 | `len(n)`、`len(min,max)`、`unique`、`nonempty` |
| 引用 / 日期 | `ref(sheet.column)`、`after(date)`、`before(date)` |
| 组合器 | `and(...)`、`or(...)`、`not(...)` |
| 指令 | `pad` / `pad(value)`（仅固定长度数组） |

组合器可任意嵌套：

```text
string:or(and(email,len(6,64)), and(phone,len(11,15)))
string:not(or(regex(^test_), oneof("admin","root")))
```

- 顶层 `,` 就是 AND，`and(...)` 可省略；嵌套在 `or(...)` 内时仍可用 `and(...)` 分组。
- 组合器作用于数组时按整体计算；直接写在规则里的标量校验器才会逐元素执行。
- `not(a,b,c)` = `NOT (a AND b AND c)`。
- 字符串参数含 `,` `;` `:` `(` `)` 时用引号包裹。

### pad（固定数组补位）

`pad` 只适用于 `[N]T`：

```text
[4]int32:pad,gte(0)     # [1,2] -> [1,2,0,0]
[3]int:pad(7)           # [1]   -> [1,7,7]
```

不足补元素零值（或指定值），超过长度报错，补位后再执行其余校验。

### ref（外键存在性）

`ref(package.column)` 校验当前值存在于指定输出包（sheet 对应的 JSON 文件）的指定列：

```text
int:ref(items.id)          # 目标 sheet 的输出名为 items
$[*][1::3]:string:ref(item.id)
```

`ref(...)` 的第一段是**该 sheet 最终输出 JSON 的文件名**（不含 `.json`），由 `--name-tpl` 计算，不使用 Excel sheet 名。例如默认模板 `{file}_{sheet}` 时写 `ref(data_item.id)`；用 `--name-tpl '{sheet}'` 时才写 `ref(item.id)`。同一次命令传入的所有文件会一起参与 ref 查找，因此不同 xlsx 之间也可以互相引用。被引用列会按自身类型转换后比较。

### types.yaml 复用

推荐 YAML：

```yaml
types:
  itemId: "int:gt(0),ref(items.id)"
  items: "$[*][1::3]:string:ref(item.id)"
```

也支持列表形式：

```yaml
types:
  - name: itemId
    format: "int:gt(0),ref(items.id)"
```

类型行直接写 `itemId` / `items` 即可，和内置类型一样使用；`#itemId` 也兼容。加载顺序：

1. `--types <path>` 显式指定
2. 输入文件同目录的 `types.yaml` / `types.yml`
3. 当前目录的 `types.yaml` / `types.yml`

命名类型必须作为整条注解使用，并检测循环引用。名字不能和内置类型冲突（如 `int`、`list`）。

---

## 特性

- **并行处理**：文件并发处理（上限 `2 × CPU 核心数`）
- **前缀排除**：使用 `-x`（前缀匹配）跳过 sheet 和列；可重复使用
- **强类型注解**：Go 风格类型、JSONPath 选择器、校验器、条件守卫、`pad`、`ref`、`types.yaml` 复用（`-t`、`-r`、`--types`）
- **智能单元格解析**：单元格中的 JSON 字符串自动反序列化（`--parse-cell-json`）
- **全字符串模式**：无类型列输出原始字符串（`--stringify-all`）
- **字段名小写**：`-l` 转小写
- **数值去零**：`88.0` → `88`，`3.14` → `3.14`
- **浮点精度修正**：`15699.999999999998` 自动归整为 `15700`
- **空值默认值**：按列类型输出零值（数值→`0`，字符串→`""`，布尔→`false`，复合类型→`null`）
- **编码**：默认 UTF-8 无 BOM，支持 `utf8-bom`、`gbk`、`gb18030`
- **.xls 支持**：通过 [grate](https://github.com/pbnjay/grate) 读取旧版 BIFF / WPS 文件

---

## 项目结构

```text
cmd/excel2json        CLI 入口
internal/core         Options、Sheet、命名工具
internal/typing       类型解析与值转换
internal/jsonpath     RFC 9535 JSONPath 子集
internal/schema       注解、校验器、types.yaml
internal/excel        .xlsx / .xls 读取
internal/convert      转换引擎与 ref
internal/export       JSON / CSV 输出
internal/app          CLI 与运行编排
```

## 开发

```bash
make build      # 构建 ./excel2json
make test       # go test ./...
make race       # go test -race ./...
make cover      # 覆盖率
make vet        # go vet ./...
make lint       # golangci-lint
```

依赖：Go 1.25+。

---

## 依赖

- [excelize](https://github.com/xuri/excelize) — XLSX 读取
- [grate](https://github.com/pbnjay/grate) — XLS (BIFF) 读取
- [cobra](https://github.com/spf13/cobra) — CLI 框架
- [errgroup](https://pkg.go.dev/golang.org/x/sync/errgroup) — 并发控制
- [yaml.v3](https://gopkg.in/yaml.v3) — types.yaml 解析

---

## 许可证

MIT，见 [LICENSE](LICENSE)。
