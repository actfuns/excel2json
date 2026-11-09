package convert

import (
	"fmt"
	"strings"
	"sync"

	"github.com/actfuns/excel2json/internal/core"
	"github.com/actfuns/excel2json/internal/schema"
	"github.com/actfuns/excel2json/internal/typing"
)

// Engine is the core conversion engine shared by JSON and CSV exporters.
type Engine struct {
	opts      core.Options
	typeCache sync.Map     // map[string]cachedType
	specCache sync.Map     // map[string]cachedSpec
	refCache  sync.Map     // map[string]refSetResult
	sheets    []core.Sheet // all sheets in the workbook, used by ref(...)
	conv      *typing.Converter
	runner    *schema.Runner
}

// New returns an Engine without workbook context. ref(...) checks
// require NewWithSheets.
func New(opts core.Options) *Engine {
	if opts.DateFormat == "" {
		opts.DateFormat = "yyyy/MM/dd"
	}
	e := &Engine{
		opts: opts,
		conv: &typing.Converter{DateFormat: opts.DateFormat, ParseCellJSON: opts.ParseCellJSON, StringifyAll: opts.StringifyAll},
	}
	e.runner = &schema.Runner{Ref: e}
	return e
}

// NewWithSheets returns an Engine that can resolve ref(...) checks
// against every sheet in the workbook.
func NewWithSheets(opts core.Options, sheets []core.Sheet) *Engine {
	e := New(opts)
	e.sheets = sheets
	return e
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

// FilterSheets returns non-empty sheets not matching --exclude prefixes.
func (e *Engine) FilterSheets(sheets []core.Sheet) []core.Sheet {
	var out []core.Sheet
	for _, s := range sheets {
		if e.skipSheet(s.Name) {
			continue
		}
		if len(s.Headers) == 0 || len(s.Rows) == 0 {
			continue
		}
		out = append(out, s)
	}
	return out
}

// ConvertSheet returns the sheet as a dict (keyed by first column) or an array, depending on ListOutput.
func (e *Engine) ConvertSheet(sheet core.Sheet) (any, error) {
	if e.opts.ListOutput {
		return e.toArray(sheet)
	}
	return e.toDict(sheet)
}

// ConvertRow transforms a single data row into a map keyed by column name.
func (e *Engine) ConvertRow(sheet core.Sheet, rowIdx int, row []string) (map[string]any, error) {
	out := make(map[string]any)

	for ci, header := range sheet.Headers {
		if e.skipColumn(header) {
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

		var required *bool
		if ci < len(sheet.Required) {
			required = sheet.Required[ci]
		}

		// Empty cell with no type annotation → skip the field entirely
		if cell == "" && colType == "" && required == nil {
			continue
		}

		val, err := e.convertCellMeta(colType, cell, required)
		if err != nil {
			return nil, fmt.Errorf("sheet [%s] data row %d, column %q: %w",
				sheet.Name, rowIdx+1, header, err)
		}

		name := header
		if e.opts.Lowercase {
			name = strings.ToLower(name)
		}
		if name == "" {
			name = fmt.Sprintf("col_%d", ci)
		}

		out[name] = val
	}
	return out, nil
}

// CSVColumns returns column metadata for the CSV writer (filtered by exclude_prefix, with optional lowercase).
func (e *Engine) CSVColumns(sheet core.Sheet) []CSVColumn {
	var cols []CSVColumn
	for ci, header := range sheet.Headers {
		if e.skipColumn(header) {
			continue
		}
		name := header
		if e.opts.Lowercase {
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

// skipSheet returns true when the sheet name starts with any exclude prefix.
func (e *Engine) skipSheet(name string) bool {
	for _, p := range e.opts.ExcludePrefixes {
		if p != "" && strings.HasPrefix(name, p) {
			return true
		}
	}
	return false
}

// skipColumn returns true when a column name starts with any exclude prefix.
func (e *Engine) skipColumn(header string) bool {
	for _, p := range e.opts.ExcludePrefixes {
		if p != "" && strings.HasPrefix(header, p) {
			return true
		}
	}
	return false
}

// toArray produces []any from every data row.
func (e *Engine) toArray(sheet core.Sheet) ([]any, error) {
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
func (e *Engine) toDict(sheet core.Sheet) (map[string]any, error) {
	// Determine key column: --key flag overrides the default (first column).
	keyName := ""
	if e.opts.KeyColumn != "" {
		target, _ := core.SplitFieldMarker(e.opts.KeyColumn)
		if e.opts.Lowercase {
			target = strings.ToLower(target)
		}
		keyName = ""
		for _, header := range sheet.Headers {
			name, _ := core.SplitFieldMarker(header)
			if e.opts.Lowercase {
				name = strings.ToLower(name)
			}
			if strings.EqualFold(name, target) && !e.skipColumn(header) {
				keyName = name
				break
			}
		}
		if keyName == "" {
			return nil, fmt.Errorf("key column %q not found or excluded", e.opts.KeyColumn)
		}
	} else if len(sheet.Headers) > 0 {
		keyName = sheet.Headers[0]
		if e.opts.Lowercase {
			keyName = strings.ToLower(keyName)
		}
		if e.skipColumn(keyName) {
			keyName = ""
		}
	}

	out := make(map[string]any, len(sheet.Rows))
	for ri, row := range sheet.Rows {
		m, err := e.ConvertRow(sheet, ri, row)
		if err != nil {
			return nil, err
		}
		// Use the first column's converted value as the key, so it's consistent
		// with the value inside the row (handles floating-point cleanup naturally).
		id := fmt.Sprintf("row_%d", ri)
		if keyName != "" {
			if v, ok := m[keyName]; ok {
				id = fmt.Sprintf("%v", v)
			} else if e.opts.UniqueKeys {
				return nil, fmt.Errorf("key column %q is empty or excluded (sheet %q row %d)", keyName, sheet.Name, ri+1)
			}
		}
		if e.opts.UniqueKeys {
			if _, exists := out[id]; exists {
				return nil, fmt.Errorf("duplicate key %q (sheet %q row %d)", id, sheet.Name, ri+1)
			}
		}
		out[id] = m
	}
	return out, nil
}

// convertCell converts and validates a raw cell string according to its type annotation.
// If colType is empty, falls back to legacy automatic behavior (number detection, ParseCellJSON).
func (e *Engine) convertCell(colType, cell string) (any, error) {
	return e.convertCellMeta(colType, cell, nil)
}

// convertCellMeta is convertCell with an optional field-level nullability
// override taken from the header name (! or ?).
func (e *Engine) convertCellMeta(colType, cell string, required *bool) (any, error) {
	if colType == "" {
		if cell == "" {
			if required != nil && *required {
				return nil, fmt.Errorf("field cannot be empty")
			}
			return "", nil
		}
		return e.conv.LegacyConvert(cell)
	}

	spec, err := e.parseSpec(colType)
	if err != nil {
		return nil, err
	}

	// Preserve the original fast path for plain type annotations. Type-level
	// !/? markers are stripped by the annotation parser, so emptiness here is
	// controlled only by the header marker or --required-by-default.
	if schema.IsPlain(spec) {
		typ := spec.Rules[0].Type
		if cell == "" {
			req := e.opts.RequireByDefault
			if required != nil {
				req = *required
			}
			if req {
				return nil, fmt.Errorf("field cannot be empty (type %s)", typ)
			}
			return typing.ZeroValue(typ), nil
		}
		return e.convertTyped(typ.String(), cell)
	}

	return e.convertFieldSpec(spec, cell, required)
}
