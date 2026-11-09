# excel2json

> **中文文档请见 [README_zh.md](README_zh.md)**

Convert Excel (`.xls`/`.xlsx`) files to **JSON** and **CSV**.

A Go port of the [excel2json](https://github.com/neil3d/excel2json) C# tool with additional CSV export support.

---

## Install

```bash
go install github.com/actfuns/excel2json@latest
```

Or build from source:

```bash
git clone https://github.com/actfuns/excel2json.git
cd excel2json
go build -o excel2json .
```

---

## Usage

```bash
excel2json [options] <inputs...> [-o <path>] [-f json|csv]
```

Inputs can be file paths or directories (`.xls`/`.xlsx` files inside are collected automatically).

### Examples

```bash
# JSON to stdout
excel2json data.xlsx

# CSV to stdout
excel2json data.xlsx -f csv

# Pretty-print JSON
excel2json data.xlsx -p

# Write to directory (each sheet → separate file)
excel2json data.xlsx -o ./output/

# Merge all sheets into one file
excel2json data.xlsx -o out.json -m

# Use custom column as dict key
excel2json data.xlsx -k cs_id

# Skip rows before data
excel2json data.xlsx -s 2

# Exclude sheets & columns starting with prefix
excel2json data.xlsx -x S_ -x cs_

# Custom output filenames
excel2json data.xlsx --name file           # just filename
excel2json data.xlsx --name sheet          # just sheet name
excel2json data.xlsx --name-tpl 'my_{file}_{sheet}'   # custom template

# Strongly typed columns (type row = row 2, data starts at row 3)
excel2json data.xlsx -t 1 -s 2 --required
```

### Options

| Short | Long | Default | Description |
|---|---|---|---|
| | `inputs...` | **required** | Excel file(s) or directory containing `.xls`/`.xlsx` |
| `-o` | `--out` | `""` | Output path (file or directory) |
| `-f` | `--format` | `json` | Output format: `json` or `csv` |
| | | **Data parsing** |
| `-n` | `--name-row` | `0` | Row index for column names (0-based) |
| `-t` | `--type-row` | `-1` | Row index for type annotations (-1 to disable) |
| `-r` | `--required` | `false` | Treat headers without `!`/`?` as required (empty cells error) |
| `-s` | `--skip-rows` | `1` | Rows to skip before data (0-based) |
| `-e` | `--encoding` | `utf8-nobom` | Output file encoding |
| `-d` | `--date-format` | `yyyy/MM/dd` | Date format string |
| | | **Field processing** |
| `-l` | `--lowcase` | `false` | Convert field names to lowercase |
| `-x` | `--exclude` | `nil` | Skip sheets & columns with prefix (repeatable) |
| | `--cell-json` | `false` | Parse JSON strings inside cells |
| | `--all-string` | `false` | Convert all values to strings |
| | `--types` | `""` | types.yaml / types.yml path (auto-discovered next to the input or in cwd) |
| | | **Output format** |
| `-a` | `--array` | `false` | Export as array (default: dict keyed by first column) |
| `-m` | `--merge` | `false` | Merge all sheets into one file (default: split) |
| | `--name-tpl` | `{file}_{sheet}` | Output filename template with `{file}` `{sheet}` |
| | `--name` | `""` | Preset: `both`, `file`, `sheet` (ignored if `--name-tpl` set) |
| `-p` | `--pretty` | `false` | Pretty-print JSON with tab indentation |
| `-k` | `--key` | `""` | Column name for dict keys (default: first column) |

---

## Output formats

### JSON — Dict mode (default)

Each row becomes an object keyed by its first-column value:

```json
{
  "1": { "name": "Alice", "age": 30 },
  "2": { "name": "Bob",   "age": 25 }
}
```

### JSON — Array mode (`-a`)

Rows are output as an array:

```json
[
  { "id": 1, "name": "Alice", "age": 30 },
  { "id": 2, "name": "Bob",   "age": 25 }
]
```

### Multiple sheets

Output is automatically split — each sheet becomes its own file named after the sheet (e.g. `Sheet1.json`, `Sheet2.json`).

Use `-m` to merge all sheets into one file:

```json
{
  "Employees": { "1": { "name": "Alice" } },
  "Departments": { "1": { "name": "Engineering" } }
}
```

### CSV

- **Single sheet**: written to the specified path.
- **Multiple sheets**: creates `{stem}_{sheetName}.csv` files.

---

## Output routing

| Input(s) | `-o` omitted | `-o <file>` | `-o <dir>/` |
|---|---|---|---|
| Single file | stdout | write to file | `{dir}/{stem}.{format}` |
| Multiple files | stdout (sequentially) | treated as directory | `{dir}/{stem}.{format}` |
| Directory | stdout (sequentially) | treated as directory | same as above |

---

## Type annotations

Set `-t/--type-row` to the row containing type annotations (0-based). Cells are then converted and strictly validated.

| Category | Types |
|---|---|
| Scalars | `bool`, `string`, `int`, `int8`, `int16`, `int32`, `int64`, `uint`, `uint8`, `uint16`, `uint32`, `uint64`, `uintptr`, `float32`, `float64`, `byte`, `rune`, `any`, `interface{}` |
| Arrays / slices | `[]T`, `[N]T` |
| Maps | `map[K]V` |
| Extras | `date`, `float` (= `float64`), `object` (= `map[string]any`), `list` (= `[]any`) |

Composite values must be valid JSON in the cell, e.g. `[1,2,3]` or `{"a":"b"}`. Array length, element types, numeric ranges, map keys and map values are all validated.

`any` / `interface{}` keep the legacy auto behavior: numbers are detected automatically, and JSON is parsed only with `--cell-json`. `date` falls back to the original string when parsing fails.

### Required / optional (header)

Emptiness is controlled only by the header marker or `--required`:

| Header | Meaning | Output key |
|---|---|---|
| `id!` | required, an empty cell is an error | `id` |
| `name?` | optional, empty cells are allowed | `name` |

Unmarked headers use `-r/--required`. Empty values produce `0`, `false`, `""`, or `null` (for composite types).

### Type matching (type-level ! / ?)

A trailing `!` / `?` on the type controls **type matching**, not emptiness:

| Annotation | Meaning |
|---|---|
| `int` / `int!` | must be an int; a mismatch is an error |
| `int?` | only apply this rule when the value is an int; otherwise skip |

```text
id!                        # required field
name?                      # optional field

int:lt(1)                  # must be int and < 1
int?:lt(1)                 # validate < 1 only when the value is an int
int?:lt(1);string?:minlen(1)
```

### Validation rules

```ebnf
fieldSpec = rule { ";" rule }
rule      = [ selector ":" ] [ type [ "!" | "?" ] [ ":" checks ] ]
checks    = check { "," check }
```

| Syntax | Meaning |
|---|---|
| `,` | all checks must pass (AND) |
| `;` | multiple rules, all applied (AND; selectors may overlap) |
| `!` | the type must match (default; explicit for clarity) |
| `?` | apply this rule only when the value matches the type; otherwise skip |
| selector omitted | applies to the field itself |
| type omitted | validates each selected node as-is |

Selectors are an RFC 9535 subset: `$`, `.field`, `['field']`/`["field"]`, `[0]`, `[-1]`, `[*]`, `[start:stop:step]`, `[i,j,k]`, `..field`. Multiple matches are flattened into a single nodelist.

```text
int:gt(0),lt(100)
int?:lt(1);string?:minlen(1)
$[*][1::3]:string:ref(item.id)
$[0]:int:gte(0);$[2]:int32:gt(0)
[4]int32:pad,gte(0)
```

- Selectors only **locate** the values a rule validates; they are not output selectors.
- Output is always the whole field, rewritten in place: selected nodes are converted and written back, untouched parts stay as-is.
- Rules run in order; later overlapping assignments win.
- A matched conditional type still reports check failures; an unmatched one is skipped.
- Missing matches produce an empty nodelist (no error).
- Scalar validators apply per element on arrays; `len`/`minlen`/`maxlen`/`unique`/`nonempty` apply to the collection.

```text
input  [[1,"a",10.0],[2,"b",20.0]]
rule   $[*][2::3]:int
output [[1,"a",10],[2,"b",20]]
```

Validators: `gt`, `gte`, `lt`, `lte`, `between`, `eq`, `ne`, `positive`, `negative`, `nonzero`, `multipleof`, `len`, `minlen`, `maxlen`, `email`, `phone`, `phone(e164)`, `url`, `uuid`, `regex`, `prefix`, `suffix`, `contains`, `alpha`, `alnum`, `numeric`, `oneof`/`in`, `notoneof`/`notin`, `unique`, `nonempty`, `ref(package.column)`, `after`, `before`, plus the nestable combinators `and(...)`, `or(...)`, `not(...)`, and the `pad` directive for fixed arrays.

`ref(package.column)` uses the final output JSON filename stem (computed from `--name-tpl`) as `package`, with `.json` optional; it falls back to the sheet name. With `--merge`, the sheet name is used.

```text
string:or(and(email,len(6,64)), and(phone,len(11,15)))
string:not(or(regex(^test_), oneof("admin","root")))
```

`not(a,b,c)` means `NOT (a AND b AND c)`. String arguments containing `,` `;` `:` `(` `)` must be quoted. `pad` only applies to `[N]T` and fills missing elements with the element zero value or an explicit value before validation.

### types.yaml

YAML is preferred:

```yaml
types:
  itemId: "int:gt(0),ref(items.id)"
  items: "$[*][1::3]:string:ref(item.id)"
```

List form is also accepted:

```yaml
types:
  - name: itemId
    format: "int:gt(0),ref(items.id)"
```

Use `itemId` / `items` in the type row, just like built-in types; `#itemId` is also accepted. Load order: `--types <path>`, then `types.yaml`/`types.yml` next to the input file, then the current directory. Named types must be the whole annotation; cycles are detected. Names may not collide with built-in types such as `int` or `list`.

---

## Features

- **Parallel processing**: files are processed concurrently (up to `2 × CPU cores`)
- **Prefix exclusion**: use `-x` (prefix match) to skip sheets & columns; repeatable
- **Strong typing**: Go-style types, JSONPath selectors, validators, conditional types, `pad`, `ref`, and `types.yaml` reuse (`-t`, `-r`, `--types`)
- **Smart cell parsing**: JSON strings inside cells are deserialised in-place (`--cell-json`)
- **All-string mode**: disable all parsing, output raw strings (`--all-string`)
- **Field name lowercasing**: `--lowcase` for case-insensitive consumers
- **Number de-dotting**: `88.0` → `88`, `3.14` → `3.14`
- **Floating-point precision fix**: values like `15699.999999999998` auto-rounded to `15700`
- **Empty-cell defaults**: empty cells use the inferred type default (`0` for numbers, `""` for strings)
- **Encoding**: UTF-8 without BOM (default)
- **.xls support**: old BIFF / WPS files are read via [grate](https://github.com/pbnjay/grate), no external tools required

---

## Dependencies

- [excelize](https://github.com/xuri/excelize) — XLSX reading
- [grate](https://github.com/pbnjay/grate) — XLS (BIFF) reading
- [cobra](https://github.com/spf13/cobra) — CLI framework
- [errgroup](https://pkg.go.dev/golang.org/x/sync/errgroup) — concurrent processing

---

## License

MIT
