package main

import (
	"path/filepath"
	"strings"
)

// Supported type annotations for the type row feature.
const (
	TypeInt    = "int"
	TypeFloat  = "float"
	TypeBool   = "bool"
	TypeString = "string"
	TypeDate   = "date"
	TypeAny    = "any"
	TypeObject = "object"
	TypeList   = "list"
)

// typeDefault returns the zero/null default value for a given type annotation.
// When a typed cell is empty, this value is emitted instead of omitting the field.
func typeDefault(colType string) any {
	switch colType {
	case TypeInt:
		return int64(0)
	case TypeFloat:
		return float64(0)
	case TypeBool:
		return false
	case TypeString, TypeDate:
		return ""
	case TypeAny, TypeObject, TypeList:
		return nil
	default:
		return ""
	}
}

// Options maps CLI flags.
type Options struct {
	// Input / output
	OutputPath string // -o --out  output path (file for single input, dir for multiple)

	// Data parsing
	NameRow    int    // --name-row     row index for column names (0-based), default 0
	TypeRow    int    // --type-row     row index for column types (0-based), -1 to disable
	HeaderRows int    // --header       #header rows between name row and data; also implies dataStart = header
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
	Types   []string   // column type annotations (empty if no type row)
	Rows    [][]string // data rows
}

// fileStem returns the filename without extension.
func fileStem(path string) string {
	name := filepath.Base(path)
	if i := strings.LastIndex(name, "."); i > 0 {
		return name[:i]
	}
	return name
}
