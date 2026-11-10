package core

import (
	"strings"
	"unicode"
)

// Sheet holds the parsed contents of a single worksheet.
type Sheet struct {
	Name       string
	RawName    string     // original sheet name including @tags
	Tags       []string   // sheet-level tags from the @tag suffix
	Headers    []string   // column names with !/? markers and @tags stripped
	ColumnTags [][]string // per-column tags from the @tag suffix
	Types      []string   // column type annotations
	Required   []*bool    // per-column nullability; nil means unset
	Rows       [][]string // data rows
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

// SplitTagSuffix extracts a trailing @tag1,tag2 suffix from a sheet or column
// name. Tags may contain letters, digits, underscores and hyphens. The suffix
// is only recognised when every comma-separated token is non-empty and valid.
func SplitTagSuffix(name string) (string, []string, bool) {
	i := strings.LastIndexByte(name, '@')
	if i <= 0 || i == len(name)-1 {
		return name, nil, false
	}
	parts := strings.Split(name[i+1:], ",")
	tags := make([]string, 0, len(parts))
	for _, part := range parts {
		tag := strings.TrimSpace(part)
		if !validTag(tag) {
			return name, nil, false
		}
		tags = append(tags, strings.ToLower(tag))
	}
	return name[:i], tags, true
}

func validTag(tag string) bool {
	if tag == "" {
		return false
	}
	for _, r := range tag {
		if !(unicode.IsLetter(r) || unicode.IsDigit(r) || r == '_' || r == '-') {
			return false
		}
	}
	return true
}

// SplitHeader parses a column header. When stripMarkers is true, trailing !/?
// required markers are removed as well. Tags and markers may appear in either
// order, so reward@c! and reward!@c both yield name reward, tags [c].
func SplitHeader(name string, stripMarkers bool) (string, *bool, []string) {
	base := strings.TrimSpace(name)
	var required *bool
	var tags []string
	for i := 0; i < 3; i++ {
		changed := false
		if stripMarkers {
			if next, req := SplitFieldMarker(base); next != base {
				base = next
				if required == nil {
					required = req
				}
				changed = true
			}
		}
		if next, suffix, ok := SplitTagSuffix(base); ok {
			base = next
			tags = append(tags, suffix...)
			changed = true
		}
		if !changed {
			break
		}
	}
	return strings.TrimSpace(base), required, tags
}

// TagMatches reports whether a tagged item is selected for target. An empty
// target selects everything; an item without tags is treated as common and
// also matches every target.
func TagMatches(tags []string, target string) bool {
	if target == "" || len(tags) == 0 {
		return true
	}
	for _, tag := range tags {
		if strings.EqualFold(tag, target) {
			return true
		}
	}
	return false
}
