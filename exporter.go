package main

import (
	"encoding/json"
	"fmt"
	"math"
	"strconv"
	"strings"
)

// Exporter is the core conversion engine shared by JSON and CSV exporters.
type Exporter struct {
	opts Options
}

// NewExporter returns an Exporter ready for use.
func NewExporter(opts Options) *Exporter {
	return &Exporter{opts: opts}
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
func (e *Exporter) ConvertSheet(sheet Sheet) (any, error) {
	if e.opts.ExportArray {
		return e.toArray(sheet)
	}
	return e.toDict(sheet)
}

// ConvertRow transforms a single data row into a map keyed by column name.
func (e *Exporter) ConvertRow(sheet Sheet, rowIdx int, row []string) (map[string]any, error) {
	out := make(map[string]any)
	idx := 0

	for ci, header := range sheet.Headers {
		if e.isExcluded(header) {
			continue
		}

		cell := ""
		if ci < len(row) {
			cell = strings.TrimSpace(row[ci])
		}

		colType := ""
		if ci < len(sheet.Types) {
			colType = sheet.Types[ci]
		}

		// Empty cell with no type annotation → skip the field entirely
		if cell == "" && colType == "" {
			continue
		}

		val, err := e.convertCell(colType, cell)
		if err != nil {
			return nil, fmt.Errorf("sheet [%s] data row %d, column %q: %w",
				sheet.Name, rowIdx+1, header, err)
		}

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
	return out, nil
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
func (e *Exporter) toArray(sheet Sheet) ([]any, error) {
	out := make([]any, 0, len(sheet.Rows))
	for ri, row := range sheet.Rows {
		m, err := e.ConvertRow(sheet, ri, row)
		if err != nil {
			return nil, err
		}
		out = append(out, m)
	}
	return out, nil
}

// toDict produces map[string]any keyed by the first column's value.
func (e *Exporter) toDict(sheet Sheet) (map[string]any, error) {
	out := make(map[string]any, len(sheet.Rows))
	for ri, row := range sheet.Rows {
		id := ""
		if len(row) > 0 {
			id = strings.TrimSpace(row[0])
		}
		if id == "" {
			id = fmt.Sprintf("row_%d", ri)
		}
		m, err := e.ConvertRow(sheet, ri, row)
		if err != nil {
			return nil, err
		}
		out[id] = m
	}
	return out, nil
}

// convertCell converts and validates a raw cell string according to its type annotation.
// If colType is empty, falls back to legacy automatic behavior (number detection, CellJSON).
func (e *Exporter) convertCell(colType, cell string) (any, error) {
	if cell == "" {
		return typeDefault(colType), nil
	}

	// --- With type annotation → strict typed conversion ---
	if colType != "" {
		switch colType {
		case TypeInt:
			v, err := strconv.ParseInt(strings.TrimSpace(cell), 10, 64)
			if err != nil {
				return nil, fmt.Errorf("expected int, got %q", cell)
			}
			return v, nil

		case TypeFloat:
			v, err := strconv.ParseFloat(strings.TrimSpace(cell), 64)
			if err != nil {
				return nil, fmt.Errorf("expected float, got %q", cell)
			}
			return v, nil

		case TypeBool:
			v, err := strconv.ParseBool(strings.ToLower(strings.TrimSpace(cell)))
			if err != nil {
				return nil, fmt.Errorf("expected bool, got %q", cell)
			}
			return v, nil

		case TypeDate:
			return cell, nil

		case TypeAny:
			return e.legacyConvert(cell)

		case TypeObject:
			trimmed := strings.TrimSpace(cell)
			if trimmed == "" || trimmed[0] != '{' {
				return nil, fmt.Errorf("expected JSON object, got %q", cell)
			}
			var parsed any
			if json.Unmarshal([]byte(trimmed), &parsed) != nil || parsed == nil {
				return nil, fmt.Errorf("expected valid JSON object, got %q", cell)
			}
			return parsed, nil

		case TypeList:
			trimmed := strings.TrimSpace(cell)
			if trimmed == "" || trimmed[0] != '[' {
				return nil, fmt.Errorf("expected JSON array, got %q", cell)
			}
			var parsed any
			if json.Unmarshal([]byte(trimmed), &parsed) != nil || parsed == nil {
				return nil, fmt.Errorf("expected valid JSON array, got %q", cell)
			}
			return parsed, nil

		case TypeString:
			return cell, nil

		default:
			// Unknown type annotation → treat as string, no error
			return cell, nil
		}
	}

	// --- No type annotation → legacy automatic behavior ---
	return e.legacyConvert(cell)
}

// legacyConvert applies CellJSON and number detection to a raw cell string.
func (e *Exporter) legacyConvert(cell string) (any, error) {
	// cell_json: try to deserialize JSON objects / arrays in-place
	if e.opts.CellJSON {
		if cell != "" && (cell[0] == '{' || cell[0] == '[') {
			var parsed any
			if json.Unmarshal([]byte(cell), &parsed) == nil && parsed != nil {
				if e.opts.AllString {
					if b, err := json.Marshal(parsed); err == nil {
						return string(b), nil
					}
				}
				return parsed, nil
			}
		}
	}

	// number parsing: 3.0 → 3 (strip redundant .0), then all_string
	if f, err := strconv.ParseFloat(cell, 64); err == nil {
		if i := int64(f); f == float64(i) && !math.IsInf(f, 0) {
			if e.opts.AllString {
				return strconv.FormatInt(i, 10), nil
			}
			return i, nil
		}
		if e.opts.AllString {
			return strconv.FormatFloat(f, 'f', -1, 64), nil
		}
		return f, nil
	}

	return cell, nil
}
