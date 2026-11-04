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
| `-r` | `--required` | `false` | Treat types without `?`/`!` as required (empty cells error) |
| `-s` | `--skip-rows` | `1` | Rows to skip before data (0-based) |
| `-e` | `--encoding` | `utf8-nobom` | Output file encoding |
| `-d` | `--date-format` | `yyyy/MM/dd` | Date format string |
| | | **Field processing** |
| `-l` | `--lowcase` | `false` | Convert field names to lowercase |
| `-x` | `--exclude` | `nil` | Skip sheets & columns with prefix (repeatable) |
| | `--cell-json` | `false` | Parse JSON strings inside cells |
| | `--all-string` | `false` | Convert all values to strings |
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

### Required / optional

| Annotation | Meaning |
|---|---|
| `T` | Uses the global default (`--required`) |
| `T!` | Required: an empty cell is an error |
| `T?` | Optional: an empty cell is allowed |

Unmarked types are optional by default. Pass `-r/--required` to make them required; explicit `?` and `!` always take precedence. Markers apply to the whole column type, not to nested elements.

Empty optional cells produce `0`, `false`, `""`, or `null` (for composite types); `?` only relaxes the empty check and does not change that output. Unknown types and invalid values produce an error with sheet, row and column context.

```text
row 1 (names):   id    tags      attrs
row 2 (types):   int   [3]int!   map[string]string?
row 3 (data):    1     [1,2,3]   {"a":"b"}
```

---

## Features

- **Parallel processing**: files are processed concurrently (up to `2 × CPU cores`)
- **Prefix exclusion**: use `-x` (prefix match) to skip sheets & columns; repeatable
- **Strong typing**: Go-style type annotations with arrays, slices, maps, numeric validation and required markers (`-t`, `-r`)
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
