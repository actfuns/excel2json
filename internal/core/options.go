// Package core defines workbook options, sheet data and naming helpers.
package core

// Options controls input parsing, conversion and output.
type Options struct {
	// Input / output.
	OutputPath string // -o, --out
	Format     string // -f, --format: json (default) or csv

	// Data parsing.
	HeaderRow  int    // row index for column names (0-based)
	TypeRow    int    // row index for type annotations (-1 to disable)
	DataRow    int    // 0-based index of the first data row (-1 = auto)
	Encoding   string // output file encoding
	DateFormat string // date format string

	// Field processing.
	Lowercase       bool     // convert field names to lowercase
	ExcludePrefixes []string // skip sheets/columns with these prefixes
	Tag             string   // only export sheets/columns tagged with this value
	StrictTags      bool     // with Tag, error on untagged exported sheets/columns
	Verbose         bool     // log skipped sheets/columns and files
	ParseCellJSON   bool     // parse JSON strings inside cells
	StringifyAll    bool     // convert all values to strings

	// Nullability. RequireByDefault treats fields without a header !/? marker
	// as required (an empty cell is an error).
	RequireByDefault bool

	// Output format.
	ListOutput  bool   // export as a list/array instead of a dict
	MergeSheets bool   // merge all sheets into one file (default: split)
	NameTmpl    string // output filename template ({file} {sheet})
	Pretty      bool   // pretty-print JSON
	KeyColumn   string // column name for dict keys (default: first column)
	UniqueKeys  bool   // error on duplicate/empty dict keys

	// Type reuse.
	TypesPath string            // path to types.yaml/types.yml (auto-discovered if empty)
	TypeDefs  map[string]string // loaded named type definitions

	// Output identity. SourceStem is the input file stem and is used to
	// compute ref(...) package names.
	SourceStem string
}
