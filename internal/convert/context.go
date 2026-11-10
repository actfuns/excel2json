package convert

import (
	"github.com/actfuns/excel2json/internal/core"
)

// SheetContext describes one sheet available to ref(...) checks. OutputName is
// the final output filename stem used to resolve references across workbooks.
type SheetContext struct {
	Sheet      core.Sheet
	OutputName string
	TypeDefs   map[string]string
}

// SheetsContext builds a context from sheets without custom output names.
func SheetsContext(opts core.Options, sheets []core.Sheet) []SheetContext {
	return sheetContexts(opts, sheets)
}

func sheetContexts(opts core.Options, sheets []core.Sheet) []SheetContext {
	out := make([]SheetContext, len(sheets))
	for i, sheet := range sheets {
		out[i] = SheetContext{Sheet: sheet, TypeDefs: opts.TypeDefs}
	}
	return out
}

// NewWithContext returns an Engine that can resolve ref(...) across every
// sheet in context, including sheets from other input files.
func NewWithContext(opts core.Options, context []SheetContext) *Engine {
	e := New(opts)
	e.sheets = context
	return e
}
