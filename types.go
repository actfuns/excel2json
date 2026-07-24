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

	// Output format
	ExportArray bool   // -a --array    export as array instead of dict
	Merge       bool   // -m --merge   merge all sheets into one file (default: split)
	NameTmpl    string // --name-tpl   output filename template {file} {sheet} (default: {file}_{sheet})
	Name        string // --name       preset: both, file, sheet
	Pretty      bool   // -p --pretty   pretty-print JSON
	KeyColumn   string // -k --key      column name for dict keys (default: first column)
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
