# excel2json

> **中文文档请见 [README_zh.md](README_zh.md)**

Convert Excel (`.xls`/`.xlsx`) files to **JSON** and **CSV**.

A Go port of the [excel2json](https://github.com/neil3d/excel2json) C# tool with additional CSV export support.

---

## Install

```bash
go install github.com/actfuns/excel2json/cmd/excel2json@latest
```

Requires Go 1.25.10 or newer.

Or build from source:

```bash
git clone https://github.com/actfuns/excel2json.git
cd excel2json
go build -o excel2json ./cmd/excel2json   # or: make build
```

---

## Usage

```bash
excel2json [options] <inputs...> [-o <path>] [-f json|csv]
excel2json --version
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
excel2json data.xlsx -o out.json

# Use custom column as dict key
excel2json data.xlsx -k cs_id

# Start reading data at row index 2 (-1 = auto)
excel2json data.xlsx --data-row 2

# Exclude sheets & columns starting with prefix
excel2json data.xlsx -x S_ -x cs_

# Custom output filenames
excel2json data.xlsx --name-tpl 'my_{file}_{sheet}'   # custom template

# Strongly typed columns (type row = row 2, data starts at row 3)
excel2json data.xlsx -t 1 --required-by-default
```

### Options

| Short | Long | Default | Description |
|---|---|---|---|
| | `inputs...` | **required** | Excel file(s) or directory containing `.xls`/`.xlsx` |
| `-o` | `--out` | `""` | Output path (file or directory) |
| `-f` | `--format` | `json` | Output format: `json` or `csv` |
| | | **Data parsing** |
| `-n` | `--header-row` | `0` | Row index for column names (0-based) |
| `-t` | `--type-row` | `-1` | Row index for type annotations (-1 to disable) |
| `-r` | `--required-by-default` | `false` | Treat headers without `!`/`?` as required (empty cells error) |
| `-s` | `--data-row` | `-1` | First data row index (0-based, -1 = auto) |
| `-e` | `--encoding` | `utf8-nobom` | File output encoding: `utf8`, `utf8-nobom`, `utf8-bom`, `gbk`, `gb18030` |
| `-d` | `--date-format` | `yyyy/MM/dd` | Date format string |
| | | **Field processing** |
| `-l` | `--lowercase` | `false` | Convert field names to lowercase |
| `-x` | `--exclude` | `nil` | Drop comment sheets/columns by prefix at load time (repeatable) |
| | `--tag` | `""` | Only export sheets/columns with this tag; untagged = common |
| | `--strict-tags` | `false` | With `--tag`, error on untagged sheets/columns |
| | `--parse-cell-json` | `false` | Parse JSON strings inside untyped cells |
| | `--stringify-all` | `false` | Stringify untyped values |
| | `--types` | `""` | types.yaml / types.yml path (auto-discovered next to the input or in cwd) |
| | | **Output format** |
| | `--list` | `false` | Export JSON as a list/array (default: dict keyed by first column) |
| `-m` | `--merge-sheets` | `false` | Merge all sheets into one file (JSON object or CSV sections) |
| | `--name-tpl` | `{file}_{sheet}` | Split-output filename template with `{file}` `{sheet}` (JSON & CSV) |
| `-p` | `--pretty` | `false` | Pretty-print JSON with tab indentation |
| `-k` | `--key` | `""` | Column name for dict keys (default: first exported column) |
| | `--unique-keys` | `false` | Error on duplicate or empty dict keys (default: last wins) |

Row indices (`--header-row`, `--type-row`, `--data-row`) refer to actual 0-based sheet rows. Blank rows are preserved for locating header/type rows and skipped in the data area. `--data-row=-1` uses the row after the header/type rows.

---

### Tags (client/server fields)

Tags select output targets without changing JSON keys:

```text
reward@c        # client only
reward@s        # server only
reward@c,s      # both
reward          # no tag = common
Shop@c          # sheet-level tag
```

```bash
excel2json --tag c -x '#' data.xlsx -o dist/client/
excel2json --tag s -x '#' data.xlsx -o dist/server/
```

- `@c,s` is metadata only: JSON keys stay `reward`; sheet filenames drop the suffix (`Shop@c` → `Shop.json`)
- `-x/--exclude` drops matching sheets/columns at load time; they are not valid for output, guards or `ref(...)` (e.g. `-x '#'` for comments)
- Untagged sheets and columns are common and exported for every `--tag`
- `--strict-tags` makes an untagged exported sheet/column an error
- Dict mode keys by the first exported column
- Column base names must be unique within a sheet, case-insensitively and independent of tags (`reward@c` + `reward@s` is invalid); `-x` excluded columns do not count
- `ref(...)` validates against all loaded sheets/columns, independent of `--tag`; `-x` excluded sheets/columns are not loaded and cannot be referenced
- Output names are checked across all loaded sheets, not just the selected tag; same output name is always an error
- Sheet tags do not make output names unique: `Shop@c` + `Shop@s` still need distinct base names or a `--name-tpl` that distinguishes them
- Workbooks/sheets with no matching columns are skipped; if nothing matches at all, `--tag` reports a clear error
- Tags may contain letters, digits, `_` and `-`; a trailing `@...` suffix is reserved for tags
- Tags and required markers can be combined: `reward@c!`, `reward!@s`
- If client/server need different values, use distinct column names or a single column tagged `@c,s`
- Without `--tag`, all sheets/columns are exported; duplicate output field names still error

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

- **`-o file.csv`**: all sheets are merged into one file.
- **`-o dir/`**: each sheet gets its own file named by `--name-tpl` (default `{file}_{sheet}.csv`), including when only one sheet matches.

---

## Output routing

| Input(s) | `-o` omitted | `-o <file>` | `-o <dir>/` |
|---|---|---|---|
| Single file | stdout | write to file | split per sheet by `--name-tpl`, or `-m` → `{dir}/{stem}.{format}` |
| Multiple files | stdout (sequentially) | treated as directory | split per sheet by `--name-tpl`, or `-m` → `{dir}/{stem}.{format}` |
| Directory | stdout (sequentially) | treated as directory | same as above |

---

## Type annotations

Set `-t/--type-row` to the row containing type annotations (0-based). Cells are then converted and strictly validated.

| Category | Types |
|---|---|
| Scalars | `bool`, `string`, `int`, `int8`, `int16`, `int32`, `int64`, `uint`, `uint8`, `uint16`, `uint32`, `uint64`, `uintptr`, `float32`, `float64`, `byte`, `rune`, `any`, `interface{}` |
| Arrays / slices | `[]T`, `[N]T` |
| Maps | `map[K]V` |
| Extras | `date`, `float` (= `float64`), `object` (= `map[string]any`), `list` (= `[]any`), `null` |

Composite values must be valid JSON in the cell, e.g. `[1,2,3]` or `{"a":"b"}`. Array length, element types, numeric ranges, map keys and map values are all validated.

`any` / `interface{}` keep the legacy auto behavior: numbers are detected automatically, and JSON is parsed only with `--parse-cell-json`. `date` falls back to the original string when parsing fails.

### Required / optional (header)

Emptiness is controlled only by the header marker or `--required-by-default`:

| Header | Meaning | Output key |
|---|---|---|
| `id!` | required, an empty cell is an error | `id` |
| `name?` | optional, empty cells are allowed | `name` |

Unmarked headers use `-r/--required-by-default`. Empty values produce `0`, `false`, `""`, or `null` (for composite types).

### when(...) guards

`when(...)` at the start of a rule enables it only when the conditions match:

```text
when(type=1) effects:[]int:ref(effect.id)
when(type!=0, level>=10) int:gt(0)
when(typeof($) = [3]int) ...
when(typeof(kind.a) = int) ...
```

- `type` — raw value of another column
- `kind.a` / `kind[0]` — JSONPath on another column
- `$.a` — JSONPath on the current cell
- `$` — the current cell
- `typeof(ref) = <type expression>` — structural JSON type check
- Operators: `=`, `==`, `!=`, `>`, `>=`, `<`, `<=`; `typeof` only supports `=`, `==`, `!=`
- Conditions inside `when(...)` are AND; `;` still separates rules
- `and(...)`, `or(...)`, `not(...)` can nest; `not(a,b)` means `NOT (a AND b)`
- `type(...)` is accepted as an alias for `typeof(...)`
- Type expressions support `[3]int`, `[]T`, `map[K]V`, `[2][2]int`, etc.
- Raw non-JSON cells are treated as `string`; `TRUE`/`FALSE` are treated as booleans
- Missing column or invalid JSON when a path is requested is an error; an empty path match means the guard is false; multiple matches use any-match

```text
when(type=1) effects:[]int:ref(effect.id)
when(typeof($) = int) gte(0)
when(typeof($) = string) len(2,4)
when(type=1, typeof(kind.a) = [3]int) ...
when(or(type=1, level>=10)) ...
when(not(type=1)) ...
```

### Validation rules

```ebnf
fieldSpec  = rule { ";" rule }
rule       = [ "when(" conditions ")" ] [ selector ":" ] [ type [ ":" checks ] ]
conditions = condition { "," condition }
checks     = check { "," check }
```

| Syntax | Meaning |
|---|---|
| `,` | all checks must pass (AND) |
| `;` | multiple rules, all applied (AND; selectors may overlap) |
| `when(...)` | rule guard; all conditions must match (AND) |
| selector omitted | applies to the field itself |
| type omitted | validates each selected node as-is |

Selectors are an RFC 9535 subset: `$`, `.field`, `['field']`/`["field"]`, `[0]`, `[-1]`, `[*]`, `[start:stop:step]`, `[i,j,k]`, `..field`. Multiple matches are flattened into a single nodelist.

```text
int:gt(0),lt(100)
when(typeof($) = int) gt(0);when(typeof($) = string) len(1)
$[*][1::3]:string:ref(item.id)
$[0]:int:gte(0);$[2]:int32:gt(0)
[4]int32:pad,gte(0)
```

- Selectors only **locate** the values a rule validates; they are not output selectors.
- Output is always the whole field, rewritten in place: selected nodes are converted and written back, untouched parts stay as-is.
- Rules run in order; later overlapping assignments win.
- A skipped guard leaves that rule unapplied; if all rules are skipped the field keeps its original value.
- Missing matches produce an empty nodelist (no error).
- Scalar validators apply per element on arrays; `len`/`minlen`/`maxlen`/`unique`/`nonempty` apply to the collection.

```text
input  [[1,"a",10.0],[2,"b",20.0]]
rule   $[*][2::3]:int
output [[1,"a",10],[2,"b",20]]
```

Validators: `gt`, `gte`, `lt`, `lte`, `between`, `eq`, `ne`, `positive`, `negative`, `nonzero`, `multipleof`, `len`, `minlen`, `maxlen`, `email`, `phone`, `phone(e164)`, `url`, `uuid`, `regex`, `prefix`, `suffix`, `contains`, `alpha`, `alnum`, `numeric`, `oneof`/`in`, `notoneof`/`notin`, `unique`, `nonempty`, `ref(package.column)`, `after`, `before`, plus the nestable combinators `and(...)`, `or(...)`, `not(...)`, and the `pad` directive for fixed arrays.

`ref(package.column)` uses the target sheet's final output JSON filename stem (computed from `--name-tpl`) as `package`; `.json` is optional. The Excel sheet name is never used implicitly. With the default template `{file}_{sheet}` write `ref(data_item.id)`; with `--name-tpl '{sheet}'` write `ref(item.id)`. Refs are resolved across every file passed in the same invocation, so different workbooks can reference each other.

```text
string:or(and(email,len(6,64)), and(phone,len(11,15)))
string:not(or(regex(^test_), oneof("admin","root")))
```

Combinators applied to an array are evaluated against the array as a whole; scalar validators written directly in a rule are applied per element. `not(a,b,c)` means `NOT (a AND b AND c)`. String arguments containing `,` `;` `:` `(` `)` must be quoted. `pad` only applies to `[N]T` and fills missing elements with the element zero value or an explicit value before validation.

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
- **Strong typing**: Go-style types, JSONPath selectors, validators, condition guards, `pad`, `ref`, and `types.yaml` reuse (`-t`, `-r`, `--types`)
- **Smart cell parsing**: JSON strings inside cells are deserialised in-place (`--parse-cell-json`)
- **Stringify mode**: output untyped columns as raw strings (`--stringify-all`)
- **Field name lowercasing**: `--lowercase` for case-insensitive consumers
- **Number de-dotting**: `88.0` → `88`, `3.14` → `3.14`
- **Floating-point precision fix**: values like `15699.999999999998` auto-rounded to `15700`
- **Empty-cell defaults**: typed columns emit their zero value (`0`, `""`, `false`, `null`)
- **Encoding**: UTF-8 without BOM by default; `utf8-bom`, `gbk` and `gb18030` are supported
- **.xls support**: old BIFF / WPS files are read via [grate](https://github.com/pbnjay/grate), no external tools required

---

## Project layout

```text
cmd/excel2json        CLI entry point
internal/core         options, sheets, naming helpers
internal/typing       type parsing and value conversion
internal/jsonpath     RFC 9535 JSONPath subset
internal/schema       annotations, validators, types.yaml
internal/excel        .xlsx / .xls reader
internal/convert      conversion engine and ref checks
internal/export       JSON / CSV output
internal/app          CLI and orchestration
```

## Development

```bash
make build      # build ./excel2json
make test       # go test ./...
make race       # go test -race ./...
make cover      # coverage
make vet        # go vet ./...
make lint       # golangci-lint
```

Requires Go 1.25.10+.

---

## Dependencies

- [excelize](https://github.com/xuri/excelize) — XLSX reading
- [grate](https://github.com/pbnjay/grate) — XLS (BIFF) reading
- [cobra](https://github.com/spf13/cobra) — CLI framework
- [errgroup](https://pkg.go.dev/golang.org/x/sync/errgroup) — concurrent processing
- [yaml.v3](https://gopkg.in/yaml.v3) — types.yaml parsing

---

## License

MIT, see [LICENSE](LICENSE).
