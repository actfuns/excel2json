package main

import (
	"encoding/json"
	"fmt"
	"strings"
)

// ExportJSON serialises one or more sheets to a JSON string.
// - Single sheet (no ForceSheetName): output is the raw array/object.
// - Multiple sheets or ForceSheetName: output is {"SheetName": <value>, ...}.
func ExportJSON(sheets []Sheet, opts Options) (string, error) {
	exp := NewExporter(opts)
	valid := exp.FilterSheets(sheets)

	if len(valid) == 0 {
		return "", fmt.Errorf("no valid sheets to export")
	}

	// single sheet & no ForceSheetName → output raw
	if !opts.ForceSheetName && len(valid) == 1 {
		v, err := exp.ConvertSheet(valid[0])
		if err != nil {
			return "", err
		}
		return serializeJSON(v, opts.Pretty)
	}

	// multiple sheets or ForceSheetName → wrap by sheet name
	data := make(map[string]any, len(valid))
	for _, s := range valid {
		v, err := exp.ConvertSheet(s)
		if err != nil {
			return "", fmt.Errorf("sheet [%s]: %w", s.Name, err)
		}
		data[s.Name] = v
	}
	return serializeJSON(data, opts.Pretty)
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
