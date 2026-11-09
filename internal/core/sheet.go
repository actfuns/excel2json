package core

import "strings"

// Sheet holds the parsed contents of a single worksheet.
type Sheet struct {
	Name     string
	Headers  []string   // column names with !/? markers stripped
	Types    []string   // column type annotations
	Required []*bool    // per-column nullability; nil means unset
	Rows     [][]string // data rows
}

// ColumnIndex returns the index of the named column, or -1.
func (s Sheet) ColumnIndex(name string) int {
	for i, header := range s.Headers {
		if header == name || strings.EqualFold(header, name) {
			return i
		}
	}
	return -1
}

// SplitFieldMarker extracts a trailing ! (required) or ? (nullable) marker
// from a header name. The returned pointer is nil when no marker is present:
// true means required, false means nullable.
func SplitFieldMarker(name string) (string, *bool) {
	switch {
	case strings.HasSuffix(name, "!"):
		v := true
		return strings.TrimSuffix(name, "!"), &v
	case strings.HasSuffix(name, "?"):
		v := false
		return strings.TrimSuffix(name, "?"), &v
	default:
		return name, nil
	}
}
