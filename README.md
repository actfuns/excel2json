# excel2json

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
excel2json [options] <inputs...> [-o <path>] [--format json|csv]
```

Inputs can be file paths or directories (`.xls`/`.xlsx` files inside are collected automatically).

### Examples

```bash
# JSON to stdout
excel2json data.xlsx

# CSV to stdout
excel2json data.xlsx --format csv

# Pretty-printed JSON to stdout
excel2json data.xlsx --pretty

# Write to file
excel2json data.xlsx -o out.json

# Write to directory
excel2json data.xlsx -o ./output/ --pretty

# Multiple files to directory
excel2json a.xlsx b.xlsx -o ./output/

# Input is a directory (scans all .xls/.xlsx)
excel2json ./data/ -o ./output/
```

### Options

| Flag | Default | Description |
|---|---|---|
| `inputs...` | **required** | Excel file(s) or directory containing `.xls`/`.xlsx` |
| `-o`, `--out` | `""` | Output path (file for single input, dir for multiple) |
| `--format` | `json` | Output format: `json` or `csv` |
| `--header` | `1` | Number of header rows |
| `-c`, `--encoding` | `utf8-nobom` | Output file encoding |
| `-l`, `--lowcase` | `false` | Convert field names to lowercase |
| `-a`, `--array` | `false` | Export as array (default: dict keyed by first column) |
| `-d`, `--date` | `yyyy/MM/dd` | Date format string |
| `-s`, `--sheet` | `false` | Force sheet-name wrapping even for single sheet |
| `-x`, `--exclude_prefix` | `""` | Skip sheets/columns starting with this prefix |
| `--cell_json` | `false` | Parse JSON strings inside cells (`{...}` / `[...]`) |
| `--all_string` | `false` | Output all values as strings (disables JSON/number parsing) |
| `--pretty` | `false` | Pretty-print JSON with tab indentation |

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

Output is automatically wrapped by sheet name:

```json
{
  "Employees": { "1": { "name": "Alice", "age": 30 } },
  "Departments": { "1": { "name": "Engineering" } }
}
```

Use `-s` to force wrapping even for a single sheet.

### CSV

- **Single sheet**: written to the specified path.
- **Multiple sheets / directory output**: creates `{sheetName}.csv` / `{stem}_{sheetName}.csv` files.

---

## Output routing

| Input(s) | `-o` omitted | `-o <file>` | `-o <dir>/` |
|---|---|---|---|
| Single file | stdout | write to file | `{dir}/{stem}.{format}` |
| Multiple files | stdout (sequentially) | treated as directory | `{dir}/{stem}.{format}` |
| Directory | stdout (sequentially) | treated as directory | same as above |

---

## Features

- **Concurrent processing**: files are processed in parallel (up to `2 × CPU cores`)
- **Smart cell parsing**: JSON strings inside cells are deserialised in-place (`--cell_json`)
- **All-string mode**: disable all parsing, output raw strings (`--all_string`)
- **Field name lowercasing**: `--lowcase` for case-insensitive consumers
- **Exclude prefix**: skip entire sheets or columns whose name starts with a prefix (`-x`)
- **Number de-dotting**: `88.0` → `88`, `3.14` → `3.14`
- **Empty-cell defaults**: the first non-empty value in a column determines the default (`0` for numbers, `""` for strings)
- **Encoding**: UTF-8 without BOM (default)

---

## Dependencies

- [excelize](https://github.com/xuri/excelize) — Excel file reading
- [cobra](https://github.com/spf13/cobra) — CLI framework
- [errgroup](https://pkg.go.dev/golang.org/x/sync/errgroup) — concurrent processing

---

## License

MIT