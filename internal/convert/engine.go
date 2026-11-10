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
	typeCache sync.Map       // map[string]cachedType
	specCache sync.Map       // map[string]cachedSpec
	refCache  sync.Map       // map[string]refSetResult
	sheets    []SheetContext // all sheets available to ref(...)
	conv      *typing.Converter
	runner    *schema.Runner
}

// New returns an Engine without workbook context. ref(...) checks
// require NewWithSheets.
func New(opts core.Options) *Engine {
	if opts.DateFormat == "" {
		opts.DateFormat = "yyyy/MM/dd"
	}
	opts.Tag = strings.ToLower(strings.TrimSpace(opts.Tag))
	e := &Engine{
		opts: opts,
		conv: &typing.Converter{DateFormat: opts.DateFormat, ParseCellJSON: opts.ParseCellJSON, StringifyAll: opts.StringifyAll},
	}
	e.runner = &schema.Runner{Ref: e}
	return e
}

// NewWithSheets returns an Engine that can resolve ref(...) checks
// against the given sheets.
func NewWithSheets(opts core.Options, sheets []core.Sheet) *Engine {
	return NewWithContext(opts, sheetContexts(opts, sheets))
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

// FilterSheets returns non-empty sheets selected by --exclude and --tag.
func (e *Engine) FilterSheets(sheets []core.Sheet) ([]core.Sheet, error) {
	var out []core.Sheet
	for _, s := range sheets {
		if e.sheetExcludedByPrefix(s) {
			continue
		}
		if len(s.Headers) == 0 || len(s.Rows) == 0 {
			continue
		}
		if err := e.validateSheetColumns(s); err != nil {
			return nil, err
		}
		if !core.TagMatches(s.Tags, e.opts.Tag) {
			continue
		}
		if e.opts.Tag != "" && !e.hasSelectedColumn(s) {
			continue
		}
		if err := e.validateSheetTags(s); err != nil {
			return nil, err
		}
		out = append(out, s)
	}
	return out, nil
}

// ConvertSheet returns the sheet as a dict (keyed by first column) or an array, depending on ListOutput.
func (e *Engine) ConvertSheet(sheet core.Sheet) (any, error) {
	if err := e.validateSheetColumns(sheet); err != nil {
		return nil, err
	}
	if err := e.validateSheetTags(sheet); err != nil {
		return nil, err
	}
	if e.opts.ListOutput {
		return e.toArray(sheet)
	}
	return e.toDict(sheet)
}

// ConvertRow transforms a single data row into a map keyed by column name.
func (e *Engine) ConvertRow(sheet core.Sheet, rowIdx int, row []string) (map[string]any, error) {
	cols, err := e.selectedColumns(sheet)
	if err != nil {
		return nil, err
	}
	return e.convertRow(sheet, rowIdx, row, cols)
}

// ConvertRowWithColumns converts a row using a previously selected column plan.
func (e *Engine) ConvertRowWithColumns(sheet core.Sheet, rowIdx int, row []string, cols []CSVColumn) (map[string]any, error) {
	return e.convertRow(sheet, rowIdx, row, cols)
}

func (e *Engine) convertRow(sheet core.Sheet, rowIdx int, row []string, cols []CSVColumn) (map[string]any, error) {
	out := make(map[string]any, len(cols))

	for _, col := range cols {
		ci := col.Index
		if ci < 0 || ci >= len(sheet.Headers) {
			return nil, fmt.Errorf("invalid column index %d in sheet %q", ci, sheet.Name)
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

		val, err := e.convertCellMetaCtx(colType, cell, required, &rowContext{sheet: &sheet, row: row})
		if err != nil {
			return nil, fmt.Errorf("sheet [%s] data row %d, column %q: %w",
				sheet.Name, rowIdx+1, sheet.Headers[ci], err)
		}

		if _, exists := out[col.Name]; exists {
			return nil, fmt.Errorf("duplicate output field %q (sheet %q, data row %d)",
				col.Name, sheet.Name, rowIdx+1)
		}
		out[col.Name] = val
	}
	return out, nil
}

func (e *Engine) selectedColumns(sheet core.Sheet) ([]CSVColumn, error) {
	cols := make([]CSVColumn, 0, len(sheet.Headers))
	for ci, header := range sheet.Headers {
		selected, err := e.columnSelected(sheet, ci)
		if err != nil {
			return nil, err
		}
		if selected {
			cols = append(cols, CSVColumn{Index: ci, Name: e.fieldName(header, ci)})
		}
	}
	return cols, nil
}

// CSVColumns returns column metadata for the CSV writer.
func (e *Engine) CSVColumns(sheet core.Sheet) ([]CSVColumn, error) {
	if err := e.validateSheetColumns(sheet); err != nil {
		return nil, err
	}
	if err := e.validateSheetTags(sheet); err != nil {
		return nil, err
	}
	selected, err := e.selectedColumns(sheet)
	if err != nil {
		return nil, err
	}
	cols := make([]CSVColumn, 0, len(selected))
	seen := make(map[string]bool, len(selected))
	for _, col := range selected {
		if seen[col.Name] {
			return nil, fmt.Errorf("duplicate output field %q (sheet %q)", col.Name, sheet.Name)
		}
		seen[col.Name] = true
		cols = append(cols, col)
	}
	return cols, nil
}

// CSVColumn describes a single column for the CSV writer.
type CSVColumn struct {
	Index int
	Name  string
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

// sheetExcludedByPrefix reports whether a sheet is hard-excluded by -x.
func (e *Engine) sheetExcludedByPrefix(sheet core.Sheet) bool {
	names := []string{sheet.Name}
	if sheet.RawName != "" && sheet.RawName != sheet.Name {
		names = append(names, sheet.RawName)
	}
	for _, name := range names {
		if e.hasExcludedPrefix(name) {
			return true
		}
	}
	return false
}

// skipColumn returns true when a column name starts with any exclude prefix.
func (e *Engine) skipColumn(header string) bool {
	return e.hasExcludedPrefix(header)
}

func (e *Engine) hasExcludedPrefix(name string) bool {
	for _, p := range e.opts.ExcludePrefixes {
		if p != "" && strings.HasPrefix(name, p) {
			return true
		}
	}
	return false
}

func (e *Engine) columnSelected(sheet core.Sheet, ci int) (bool, error) {
	header := sheet.Headers[ci]
	if e.skipColumn(header) {
		return false, nil
	}
	tags := columnTags(sheet, ci)
	if !core.TagMatches(tags, e.opts.Tag) {
		return false, nil
	}
	if e.opts.StrictTags && e.opts.Tag != "" && len(tags) == 0 {
		return false, fmt.Errorf("column %q has no tags (required by --strict-tags) in sheet %q", header, sheet.Name)
	}
	return true, nil
}

func columnTags(sheet core.Sheet, ci int) []string {
	if ci >= 0 && ci < len(sheet.ColumnTags) {
		return sheet.ColumnTags[ci]
	}
	return nil
}

func (e *Engine) fieldName(header string, ci int) string {
	name := header
	if e.opts.Lowercase {
		name = strings.ToLower(name)
	}
	if name == "" {
		name = fmt.Sprintf("col_%d", ci)
	}
	return name
}

func (e *Engine) validateSheetColumns(sheet core.Sheet) error {
	seen := make(map[string]string, len(sheet.Headers))
	for ci, header := range sheet.Headers {
		if e.skipColumn(header) {
			continue
		}
		name := e.fieldName(header, ci)
		key := strings.ToLower(name)
		if prev, ok := seen[key]; ok {
			return fmt.Errorf("duplicate column %q (also %q) in sheet %q", name, prev, sheet.Name)
		}
		seen[key] = name
	}
	return nil
}

func (e *Engine) validateSheetTags(sheet core.Sheet) error {
	if e.opts.StrictTags && e.opts.Tag != "" && len(sheet.Tags) == 0 {
		return fmt.Errorf("sheet %q has no tags (required by --strict-tags)", sheet.Name)
	}
	return nil
}

// HasSelectedColumns reports whether the sheet has at least one column selected by -x/tag.
func (e *Engine) HasSelectedColumns(sheet core.Sheet) bool {
	return e.hasSelectedColumn(sheet)
}

func (e *Engine) hasSelectedColumn(sheet core.Sheet) bool {
	for ci := range sheet.Headers {
		if e.skipColumn(sheet.Headers[ci]) {
			continue
		}
		if core.TagMatches(columnTags(sheet, ci), e.opts.Tag) {
			return true
		}
	}
	return false
}

func sheetHasTags(sheet core.Sheet) bool {
	if len(sheet.Tags) > 0 {
		return true
	}
	for _, tags := range sheet.ColumnTags {
		if len(tags) > 0 {
			return true
		}
	}
	return false
}

// toArray produces []any from every data row.
func (e *Engine) toArray(sheet core.Sheet) ([]any, error) {
	cols, err := e.selectedColumns(sheet)
	if err != nil {
		return nil, err
	}
	out := make([]any, 0, len(sheet.Rows))
	for ri, row := range sheet.Rows {
		m, err := e.convertRow(sheet, ri, row, cols)
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
	cols, err := e.selectedColumns(sheet)
	if err != nil {
		return nil, err
	}

	keyName := ""
	if e.opts.KeyColumn != "" {
		target, _, _ := core.SplitHeader(e.opts.KeyColumn, true)
		if e.opts.Lowercase {
			target = strings.ToLower(target)
		}
		for _, col := range cols {
			if strings.EqualFold(col.Name, target) {
				keyName = col.Name
				break
			}
		}
		if keyName == "" {
			return nil, fmt.Errorf("key column %q not found or excluded", e.opts.KeyColumn)
		}
	} else if len(cols) > 0 {
		if e.opts.Tag != "" || sheetHasTags(sheet) {
			keyName = cols[0].Name
		} else if cols[0].Index == 0 {
			keyName = cols[0].Name
		}
	}

	out := make(map[string]any, len(sheet.Rows))
	for ri, row := range sheet.Rows {
		m, err := e.convertRow(sheet, ri, row, cols)
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
	return e.convertCellMetaCtx(colType, cell, required, nil)
}

func (e *Engine) convertCellMetaCtx(colType, cell string, required *bool, row *rowContext) (any, error) {
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

	// Preserve the original fast path for plain type annotations. Emptiness is
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

	return e.convertFieldSpec(spec, cell, required, row)
}
