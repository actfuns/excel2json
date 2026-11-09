package main

import (
	"path/filepath"
	"strings"
)

// Options maps CLI flags.
type Options struct {
	// Input / output
	OutputPath string // -o --out
	Format     string // -f --format   json (default) or csv

	// Data parsing
	NameRow    int    // -n --name-row      row index for column names (0-based)
	TypeRow    int    // -t --type-row      row index for type annotations (-1 to disable)
	SkipRows   int    // --skip   rows to skip before data (0-based)
	Encoding   string // -e --encoding      output file encoding
	DateFormat string // -d --date-format   date format string

	// Field processing
	Lowcase       bool     // -l --lowcase         convert field names to lowercase
	ExcludePrefix []string // -x --exclude   skip sheets/columns with prefix (repeatable)
	CellJSON      bool     // --cell-json          parse JSON strings inside cells
	AllString     bool     // --all-string         convert all values to strings

	// Nullability
	// Required treats types without ?/! as required (empty cell errors).
	Required bool // -r --required

	// Output format
	ExportArray bool   // -a --array    export as array instead of dict
	Merge       bool   // -m --merge   merge all sheets into one file (default: split)
	NameTmpl    string // --name-tpl   output filename template {file} {sheet} (default: {file}_{sheet})
	Name        string // --name       preset: both, file, sheet
	Pretty      bool   // -p --pretty   pretty-print JSON
	KeyColumn   string // -k --key      column name for dict keys (default: first column)

	// Type reuse
	TypesPath string            // --types   path to types.yaml/types.yml (auto-discovered if empty)
	TypeDefs  map[string]string // loaded #name -> annotation definitions (not a CLI flag)

	// Output identity
	SourceName string // input file stem, used to compute ref(...) package names (not a CLI flag)
}

// Sheet holds the parsed contents of a single worksheet.
type Sheet struct {
	Name     string
	Headers  []string   // column names from the header rows (!/? markers stripped)
	Types    []string   // column type annotations (empty if no type row)
	Required []*bool    // per-column nullability from header markers; nil = unset
	Rows     [][]string // data rows
}

// splitFieldMarker extracts a trailing ! (required) or ? (nullable) marker
// from a header name. The returned pointer is nil when no marker is present:
// true means required, false means nullable.
func splitFieldMarker(name string) (string, *bool) {
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

// fileStem returns the filename without extension.
func fileStem(path string) string {
	name := filepath.Base(path)
	if i := strings.LastIndex(name, "."); i > 0 {
		return name[:i]
	}
	return name
}
