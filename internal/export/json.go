package export

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/actfuns/excel2json/internal/convert"
	"github.com/actfuns/excel2json/internal/core"
)

// JSON serialises one or more sheets to a JSON string.
// - Single sheet: output is the raw array/object.
// - Multiple sheets: output is {"SheetName": <value>, ...}.
func JSON(sheets []core.Sheet, opts core.Options) (string, error) {
	return JSONWithContext(sheets, convert.SheetsContext(opts, sheets), opts)
}

// JSONWithContext serialises exportSheets while exposing every sheet in
// context to ref(...) checks. This lets split exports still validate foreign
// keys against sheets written to other files.
func JSONWithContext(exportSheets []core.Sheet, context []convert.SheetContext, opts core.Options) (string, error) {
	exp := convert.NewWithContext(opts, context)
	valid, err := exp.FilterSheets(exportSheets)
	if err != nil {
		return "", err
	}

	if len(valid) == 0 {
		return "", fmt.Errorf("no valid sheets to export")
	}

	if len(valid) == 1 {
		v, err := exp.ConvertSheet(valid[0])
		if err != nil {
			return "", err
		}
		return serializeJSON(v, opts.Pretty)
	}

	// multiple sheets → wrap by sheet name
	data := make(map[string]any, len(valid))
	for _, s := range valid {
		v, err := exp.ConvertSheet(s)
		if err != nil {
			return "", fmt.Errorf("sheet [%s]: %w", s.Name, err)
		}
		if _, exists := data[s.Name]; exists {
			return "", fmt.Errorf("duplicate sheet output name %q after tag filtering", s.Name)
		}
		data[s.Name] = v
	}
	return serializeJSON(data, opts.Pretty)
}

// Serialize encodes a converted value as JSON.
func Serialize(v any, pretty bool) (string, error) {
	return serializeJSON(v, pretty)
}

func serializeJSON(v any, pretty bool) (string, error) {
	var buf strings.Builder
	enc := json.NewEncoder(&buf)
	enc.SetEscapeHTML(false)
	if pretty {
		enc.SetIndent("", "\t")
	} else {
		enc.SetIndent("", "")
	}
	if err := enc.Encode(v); err != nil {
		return "", fmt.Errorf("json encode: %w", err)
	}
	return strings.TrimRight(buf.String(), "\n"), nil
}
