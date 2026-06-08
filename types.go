package main

import (
	"path/filepath"
	"strings"
)

// Options maps CLI flags.
type Options struct {
	// Input / output
	OutputPath string // -o --out  output path (file for single input, dir for multiple)

	// Data parsing
	HeaderRows int    // --header       number of header rows, default 1
	Encoding   string // -c --encoding  output encoding, default utf8-nobom
	DateFormat string // -d --date      date format, default yyyy/MM/dd

	// Field processing
	Lowcase       bool   // -l --lowcase         convert field names to lowercase
	ExcludePrefix string // -x --exclude_prefix  skip sheets/columns with this prefix
	CellJSON      bool   // --cell_json          parse JSON strings inside cells
	AllString     bool   // --all_string         convert all values to strings

	// Output format
	Format         string // --format  json (default) or csv
	ExportArray    bool   // -a --array     export as array (default: dict keyed by first column)
	ForceSheetName bool   // -s --sheet     wrap output in sheet name even for single sheet
	Pretty         bool   // --pretty       pretty-print JSON with tab indentation
}

// Sheet holds the parsed contents of a single worksheet.
type Sheet struct {
	Name    string
	Headers []string   // column names from the header rows
	Rows    [][]string // data rows (header rows excluded)
}

// fileStem returns the filename without extension.
func fileStem(path string) string {
	name := filepath.Base(path)
	if i := strings.LastIndex(name, "."); i > 0 {
		return name[:i]
	}
	return name
}
