package main

import (
	"encoding/json"
	"fmt"
	"math"
	"strconv"
	"strings"
	"sync"
	"time"
)

// Exporter is the core conversion engine shared by JSON and CSV exporters.
type Exporter struct {
	opts      Options
	typeCache sync.Map // map[string]cachedType
	specCache sync.Map // map[string]cachedFieldSpec
	refCache  sync.Map // map[string]refSetResult
	sheets    []Sheet  // all sheets in the workbook, used by ref(...)
}

// NewExporter returns an Exporter without workbook context. ref(...) checks
// require NewExporterWithSheets.
func NewExporter(opts Options) *Exporter {
	return &Exporter{opts: opts}
}

// NewExporterWithSheets returns an Exporter that can resolve ref(...) checks
// against every sheet in the workbook.
func NewExporterWithSheets(opts Options, sheets []Sheet) *Exporter {
	return &Exporter{opts: opts, sheets: sheets}
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

// FilterSheets returns non-empty sheets not matching --exclude prefixes.
func (e *Exporter) FilterSheets(sheets []Sheet) []Sheet {
	var out []Sheet
	for _, s := range sheets {
		if e.skipSheet(s.Name) {
			continue
		}
		if len(s.Headers) == 0 || len(s.Rows) == 0 {
			continue
		}
		out = append(out, s)
	}
	return out
}

// ConvertSheet returns the sheet as a dict (keyed by first column) or an array, depending on ExportArray.
func (e *Exporter) ConvertSheet(sheet Sheet) (any, error) {
	if e.opts.ExportArray {
		return e.toArray(sheet)
	}
	return e.toDict(sheet)
}

// ConvertRow transforms a single data row into a map keyed by column name.
func (e *Exporter) ConvertRow(sheet Sheet, rowIdx int, row []string) (map[string]any, error) {
	out := make(map[string]any)

	for ci, header := range sheet.Headers {
		if e.skipColumn(header) {
			continue
		}

		cell := ""
		if ci < len(row) {
			cell = strings.TrimSpace(row[ci])
		}

		colType := ""
		if ci < len(sheet.Types) {
			colType = sheet.Types[ci]
		}

		var required *bool
		if ci < len(sheet.Required) {
			required = sheet.Required[ci]
		}

		// Empty cell with no type annotation → skip the field entirely
		if cell == "" && colType == "" && required == nil {
			continue
		}

		val, err := e.convertCellMeta(colType, cell, required)
		if err != nil {
			return nil, fmt.Errorf("sheet [%s] data row %d, column %q: %w",
				sheet.Name, rowIdx+1, header, err)
		}

		name := header
		if e.opts.Lowcase {
			name = strings.ToLower(name)
		}
		if name == "" {
			name = fmt.Sprintf("col_%d", ci)
		}

		out[name] = val
	}
	return out, nil
}

// CSVColumns returns column metadata for the CSV writer (filtered by exclude_prefix, with optional lowercase).
func (e *Exporter) CSVColumns(sheet Sheet) []CSVColumn {
	var cols []CSVColumn
	for ci, header := range sheet.Headers {
		if e.skipColumn(header) {
			continue
		}
		name := header
		if e.opts.Lowcase {
			name = strings.ToLower(name)
		}
		if name == "" {
			name = fmt.Sprintf("col_%d", ci)
		}
		cols = append(cols, CSVColumn{Index: ci, Name: name})
	}
	return cols
}

// CSVColumn describes a single column for the CSV writer.
type CSVColumn struct {
	Index int
	Name  string
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

// skipSheet returns true when the sheet name starts with any exclude prefix.
func (e *Exporter) skipSheet(name string) bool {
	for _, p := range e.opts.ExcludePrefix {
		if p != "" && strings.HasPrefix(name, p) {
			return true
		}
	}
	return false
}

// skipColumn returns true when a column name starts with any exclude prefix.
func (e *Exporter) skipColumn(header string) bool {
	for _, p := range e.opts.ExcludePrefix {
		if p != "" && strings.HasPrefix(header, p) {
			return true
		}
	}
	return false
}

// toArray produces []any from every data row.
func (e *Exporter) toArray(sheet Sheet) ([]any, error) {
	out := make([]any, 0, len(sheet.Rows))
	for ri, row := range sheet.Rows {
		m, err := e.ConvertRow(sheet, ri, row)
		if err != nil {
			return nil, err
		}
		out = append(out, m)
	}
	return out, nil
}

// toDict produces map[string]any keyed by the first column's value.
func (e *Exporter) toDict(sheet Sheet) (map[string]any, error) {
	// Determine key column: --key flag overrides the default (first column).
	keyName := ""
	if e.opts.KeyColumn != "" {
		keyName, _ = splitFieldMarker(e.opts.KeyColumn)
		if e.opts.Lowcase {
			keyName = strings.ToLower(keyName)
		}
	} else if len(sheet.Headers) > 0 {
		keyName = sheet.Headers[0]
		if e.opts.Lowcase {
			keyName = strings.ToLower(keyName)
		}
		if e.skipColumn(keyName) {
			keyName = ""
		}
	}

	out := make(map[string]any, len(sheet.Rows))
	for ri, row := range sheet.Rows {
		m, err := e.ConvertRow(sheet, ri, row)
		if err != nil {
			return nil, err
		}
		// Use the first column's converted value as the key, so it's consistent
		// with the value inside the row (handles floating-point cleanup naturally).
		id := fmt.Sprintf("row_%d", ri)
		if keyName != "" {
			if v, ok := m[keyName]; ok {
				id = fmt.Sprintf("%v", v)
			}
		}
		out[id] = m
	}
	return out, nil
}

// convertCell converts and validates a raw cell string according to its type annotation.
// If colType is empty, falls back to legacy automatic behavior (number detection, CellJSON).
func (e *Exporter) convertCell(colType, cell string) (any, error) {
	return e.convertCellMeta(colType, cell, nil)
}

// convertCellMeta is convertCell with an optional field-level nullability
// override taken from the header name (! or ?).
func (e *Exporter) convertCellMeta(colType, cell string, required *bool) (any, error) {
	if colType == "" {
		if cell == "" {
			if required != nil && *required {
				return nil, fmt.Errorf("field cannot be empty")
			}
			return "", nil
		}
		return e.legacyConvert(cell)
	}

	spec, err := e.parseFieldSpecCached(colType)
	if err != nil {
		return nil, err
	}

	// Preserve the original fast path for plain type annotations. Type-level
	// !/? markers are stripped by the annotation parser, so emptiness here is
	// controlled only by the header marker or --required.
	if isPlainSpec(spec) {
		typ := spec.rules[0].typ
		if cell == "" {
			req := e.opts.Required
			if required != nil {
				req = *required
			}
			if req {
				return nil, fmt.Errorf("field cannot be empty (type %s)", typ)
			}
			return zeroValue(typ), nil
		}
		return e.convertTyped(typ.String(), cell)
	}

	return e.convertFieldSpec(spec, cell, required)
}

func isPlainSpec(spec *fieldSpec) bool {
	if spec == nil || len(spec.rules) != 1 {
		return false
	}
	rule := spec.rules[0]
	return rule.selector == nil && !rule.conditional && len(rule.checks) == 0 && rule.typ != nil
}

// legacyConvert applies CellJSON and number detection to a raw cell string.
func (e *Exporter) legacyConvert(cell string) (any, error) {
	// cell_json: try to deserialize JSON objects / arrays in-place
	if e.opts.CellJSON {
		if cell != "" && (cell[0] == '{' || cell[0] == '[') {
			var parsed any
			if json.Unmarshal([]byte(cell), &parsed) == nil && parsed != nil {
				if e.opts.AllString {
					if b, err := json.Marshal(parsed); err == nil {
						return string(b), nil
					}
				}
				return parsed, nil
			}
		}
	}

	// number parsing: 3.0 → 3 (strip redundant .0), then all_string
	if f, err := strconv.ParseFloat(cell, 64); err == nil && !math.IsNaN(f) && !math.IsInf(f, 0) {
		// Use epsilon tolerance to catch floating-point imprecision
		// (e.g. Excel may store 15700 as 15699.999999999998 internally).
		if i := int64(math.Round(f)); math.Abs(f-float64(i)) < 1e-9 {
			if e.opts.AllString {
				return strconv.FormatInt(i, 10), nil
			}
			return i, nil
		}
		if e.opts.AllString {
			return strconv.FormatFloat(f, 'f', -1, 64), nil
		}
		return f, nil
	}

	return cell, nil
}

// dateValue wraps a time.Time with the configured format so that MarshalJSON
// produces the desired date string. This pushes formatting to the serialization
// layer, analogous to C#'s JsonSerializerSettings.DateFormatString.
type dateValue struct {
	t      time.Time
	format string
}

func (d dateValue) MarshalJSON() ([]byte, error) {
	return json.Marshal(d.t.Format(d.format))
}

func (d dateValue) String() string {
	return d.t.Format(d.format)
}

// convertDate converts a raw cell value (expected to be an Excel date serial number
// or a parseable date string) into a dateValue that formats itself during JSON encoding.
func (e *Exporter) convertDate(cell string) (any, error) {
	layout := excelDateFormatToGo(e.opts.DateFormat)
	trimmed := strings.TrimSpace(cell)

	// 1) Try integer Excel serial number (e.g. "45341")
	if serial, err := strconv.ParseInt(trimmed, 10, 64); err == nil && serial >= 1 {
		return dateValue{t: excelSerialToTime(serial), format: layout}, nil
	}

	// 2) Try float Excel serial number (e.g. "45341.5" for date+time)
	if f, err := strconv.ParseFloat(trimmed, 64); err == nil && f >= 1 {
		return dateValue{t: excelSerialToFloatTime(f), format: layout}, nil
	}

	// 3) Try common date string layouts
	commonLayouts := []string{
		"2006-01-02",
		"2006/01/02",
		"01/02/2006",
		"1/2/2006",
		"2006-01-02 15:04:05",
		"2006/01/02 15:04:05",
		time.RFC3339,
	}
	for _, try := range commonLayouts {
		if t, err := time.Parse(try, trimmed); err == nil {
			return dateValue{t: t, format: layout}, nil
		}
	}

	// 4) Fallback: return as-is (will become a JSON string)
	return trimmed, nil
}

// excelSerialToTime converts an integer Excel date serial number to time.Time.
// Excel serial number 1 = January 1, 1900. Due to the Lotus 123 bug, Excel
// treats serial 60 as Feb 29, 1900 (which never existed). With the 1899-12-30
// epoch, serials below 60 therefore need a +1 day adjustment.
func excelSerialToTime(serial int64) time.Time {
	if serial < 60 {
		serial++
	}
	return time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC).AddDate(0, 0, int(serial))
}

// excelSerialToFloatTime converts a float Excel serial number to time.Time,
// preserving the fractional (time-of-day) component.
func excelSerialToFloatTime(serial float64) time.Time {
	whole := int64(serial)
	frac := serial - float64(whole)
	t := excelSerialToTime(whole)
	// frac is fraction of a day
	nanos := int64(frac * 86400 * 1e9)
	return t.Add(time.Duration(nanos))
}

// excelDateFormatToGo converts an Excel-style date format string to a Go time.Format layout.
// Examples:
//
//	yyyy/MM/dd  → 2006/01/02
//	yyyy-MM-dd  → 2006-01-02
//	MM/dd/yyyy  → 01/02/2006
//	yyyy-MM-dd HH:mm:ss → 2006-01-02 15:04:05
func excelDateFormatToGo(format string) string {
	// If it's already a Go-style layout (contains "2006" or "01" or "02"), use as-is.
	if strings.Contains(format, "2006") || strings.Contains(format, "01") || strings.Contains(format, "02") {
		return format
	}

	var buf strings.Builder
	i := 0
	for i < len(format) {
		// collect a run of letter characters
		start := i
		for i < len(format) && isDateLetter(format[i]) {
			i++
		}
		if i > start {
			buf.WriteString(replaceDateToken(format[start:i]))
		}
		if i < len(format) {
			buf.WriteByte(format[i])
			i++
		}
	}
	return buf.String()
}

func isDateLetter(c byte) bool {
	return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')
}

func replaceDateToken(token string) string {
	switch token {
	case "yyyy":
		return "2006"
	case "yy":
		return "06"
	case "MMMM":
		return "January"
	case "MMM":
		return "Jan"
	case "MM":
		return "01"
	case "M":
		return "1"
	case "dddd":
		return "Monday"
	case "ddd":
		return "Mon"
	case "dd":
		return "02"
	case "d":
		return "2"
	case "HH":
		return "15"
	case "H":
		return "15"
	case "hh":
		return "03"
	case "h":
		return "3"
	case "mm":
		return "04"
	case "m":
		return "4"
	case "ss":
		return "05"
	case "s":
		return "5"
	case "tt", "TT", "AM/PM", "am/pm":
		return "PM"
	default:
		return token
	}
}
