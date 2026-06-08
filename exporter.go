package main

import (
	"encoding/json"
	"fmt"
	"math"
	"strconv"
	"strings"
)

// Exporter is the core conversion engine shared by JSON and CSV exporters.
// It holds Options and a column-default-value cache to avoid parameter threading.
type Exporter struct {
	opts  Options
	cache map[string]any
}

// NewExporter returns an Exporter ready for use.
func NewExporter(opts Options) *Exporter {
	return &Exporter{opts: opts, cache: make(map[string]any)}
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

// FilterSheets returns non-empty sheets whose names don't start with exclude_prefix.
func (e *Exporter) FilterSheets(sheets []Sheet) []Sheet {
	var out []Sheet
	for _, s := range sheets {
		if e.opts.ExcludePrefix != "" && strings.HasPrefix(s.Name, e.opts.ExcludePrefix) {
			continue
		}
		if len(s.Headers) == 0 || len(s.Rows) == 0 {
			continue
		}
		out = append(out, s)
	}
	return out
}

// ConvertSheet returns the sheet as a dict (keyed by first column) or an array, depending on ExportArray.
func (e *Exporter) ConvertSheet(sheet Sheet) any {
	if e.opts.ExportArray {
		return e.toArray(sheet)
	}
	return e.toDict(sheet)
}

// ConvertRow transforms a single data row into a map keyed by column name.
func (e *Exporter) ConvertRow(sheet Sheet, row []string) map[string]any {
	out := make(map[string]any)
	idx := 0

	for ci, header := range sheet.Headers {
		if e.isExcluded(header) {
			continue
		}

		cell := ""
		if ci < len(row) {
			cell = row[ci]
		}

		val := e.convertCell(sheet, ci, cell)

		name := header
		if e.opts.Lowcase {
			name = strings.ToLower(name)
		}
		if name == "" {
			name = fmt.Sprintf("col_%d", idx)
		}

		out[name] = val
		idx++
	}
	return out
}

// CSVColumns returns column metadata for the CSV writer (filtered by exclude_prefix, with optional lowercase).
func (e *Exporter) CSVColumns(sheet Sheet) []CSVColumn {
	var cols []CSVColumn
	for ci, header := range sheet.Headers {
		if e.isExcluded(header) {
			continue
		}
		name := header
		if e.opts.Lowcase {
			name = strings.ToLower(name)
		}
		if name == "" {
			name = fmt.Sprintf("col_%d", ci)
		}
		cols = append(cols, CSVColumn{Index: ci, Name: name})
	}
	return cols
}

// CSVColumn describes a single column for the CSV writer.
type CSVColumn struct {
	Index int
	Name  string
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

// isExcluded returns true when the column header starts with ExcludePrefix.
func (e *Exporter) isExcluded(header string) bool {
	return e.opts.ExcludePrefix != "" && strings.HasPrefix(header, e.opts.ExcludePrefix)
}

// toArray produces []any from every data row.
func (e *Exporter) toArray(sheet Sheet) []any {
	out := make([]any, 0, len(sheet.Rows))
	for _, row := range sheet.Rows {
		out = append(out, e.ConvertRow(sheet, row))
	}
	return out
}

// toDict produces map[string]any keyed by the first column's value.
func (e *Exporter) toDict(sheet Sheet) map[string]any {
	out := make(map[string]any, len(sheet.Rows))
	for i, row := range sheet.Rows {
		id := ""
		if len(row) > 0 {
			id = strings.TrimSpace(row[0])
		}
		if id == "" {
			id = fmt.Sprintf("row_%d", i)
		}
		out[id] = e.ConvertRow(sheet, row)
	}
	return out
}

// convertCell applies all transformations to a raw cell string:
// empty→default → cell_json → number parsing (strip .0) → all_string → raw.
func (e *Exporter) convertCell(sheet Sheet, colIdx int, cell string) any {
	// Trim spaces — excelize returns raw strings including accidental spaces.
	// Matches C# ExcelDataReader which coerces space-only cells to DBNull for
	// typed columns, then getColumnDefault returns the numeric zero.
	cell = strings.TrimSpace(cell)

	if e.opts.AllString {
		return cell
	}

	if cell == "" {
		return e.columnDefault(sheet, colIdx)
	}

	// cell_json: try to deserialize JSON objects / arrays in-place
	if e.opts.CellJSON {
		if cell[0] == '{' || cell[0] == '[' {
			var parsed any
			if json.Unmarshal([]byte(cell), &parsed) == nil && parsed != nil {
				return parsed
			}
		}
	}

	// number parsing: 3.0 → 3 (strip redundant .0)
	if f, err := strconv.ParseFloat(cell, 64); err == nil {
		if i := int64(f); f == float64(i) && !math.IsInf(f, 0) {
			return i
		}
		return f
	}

	return cell
}

// columnDefault returns a zero-value for an empty cell by inspecting the first non-empty cell in the same column.
func (e *Exporter) columnDefault(sheet Sheet, colIdx int) any {
	key := sheet.Name + "@" + sheet.Headers[colIdx]
	if v, ok := e.cache[key]; ok {
		return v
	}

	result := any("")
	for _, row := range sheet.Rows {
		if colIdx >= len(row) {
			continue
		}
		val := strings.TrimSpace(row[colIdx])
		if val == "" {
			continue
		}
		if !e.opts.AllString {
			if _, err := strconv.ParseFloat(val, 64); err == nil {
				result = int64(0)
				break
			}
		}
		result = ""
		break
	}

	e.cache[key] = result
	return result
}
