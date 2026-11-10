package export

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/actfuns/excel2json/internal/convert"
	"github.com/actfuns/excel2json/internal/core"
	"github.com/actfuns/excel2json/internal/typing"
)

// CSV writes all valid sheets and returns the paths it wrote.
//   - No output path → stdout (no paths returned).
//   - --merge-sheets → one file.
//   - Single sheet → one file.
//   - Multiple sheets → split files named by --name-tpl.
func CSV(sheets []core.Sheet, opts core.Options, outPath string) ([]string, error) {
	return CSVWithContext(sheets, convert.SheetsContext(opts, sheets), opts, outPath)
}

// CSVWithContext writes sheets while resolving ref(...) against context.
func CSVWithContext(sheets []core.Sheet, context []convert.SheetContext, opts core.Options, outPath string) ([]string, error) {
	exp := convert.NewWithContext(opts, context)
	valid := exp.FilterSheets(sheets)
	if len(valid) == 0 {
		return nil, fmt.Errorf("no valid sheets to export")
	}

	if outPath == "" {
		return nil, writeCSVSeq(os.Stdout, valid, exp)
	}

	if opts.MergeSheets {
		if core.IsDirPath(outPath) {
			return nil, fmt.Errorf("--merge-sheets requires a file path, got directory %q", outPath)
		}
		f, err := os.Create(outPath)
		if err != nil {
			return nil, fmt.Errorf("create csv: %w", err)
		}
		defer f.Close()
		if err := writeCSVSeq(f, valid, exp); err != nil {
			return nil, err
		}
		return []string{outPath}, nil
	}

	if len(valid) == 1 {
		target := outPath
		if core.IsDirPath(outPath) {
			target = filepath.Join(outPath, csvOutputName(opts, valid[0].Name, "", true)+".csv")
		}
		if err := writeCSVFile(target, valid[0], exp); err != nil {
			return nil, err
		}
		return []string{target}, nil
	}

	dirMode := core.IsDirPath(outPath)
	dir := outPath
	stem := ""
	if !dirMode {
		dir = filepath.Dir(outPath)
		stem = strings.TrimSuffix(filepath.Base(outPath), ".csv")
	}
	if err := os.MkdirAll(dir, 0755); err != nil {
		return nil, fmt.Errorf("mkdir %s: %w", dir, err)
	}

	paths := make([]string, 0, len(valid))
	seen := make(map[string]bool, len(valid))
	for _, sheet := range valid {
		name := csvOutputName(opts, sheet.Name, stem, dirMode)
		path := filepath.Join(dir, name+".csv")
		if seen[path] {
			return nil, fmt.Errorf("output name collision %q (check --name-tpl)", path)
		}
		seen[path] = true
		if err := writeCSVFile(path, sheet, exp); err != nil {
			return nil, fmt.Errorf("write sheet [%s] csv: %w", sheet.Name, err)
		}
		paths = append(paths, path)
	}
	return paths, nil
}

func csvOutputName(opts core.Options, sheetName, stem string, dirMode bool) string {
	if opts.NameTmpl != "" {
		return core.SanitizeName(core.ExpandNameTemplate(opts.NameTmpl, opts.SourceStem, sheetName))
	}
	if dirMode {
		return core.SanitizeName(sheetName)
	}
	return core.SanitizeName(stem + "_" + sheetName)
}

// writeCSVFile writes a single sheet to a CSV file.
func writeCSVFile(path string, sheet core.Sheet, exp *convert.Engine) error {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("create csv: %w", err)
	}
	defer f.Close()
	return writeCSV(f, sheet, exp)
}

// writeCSVSeq writes all sheets sequentially to a writer.
// Multiple sheets are separated by comment lines.
func writeCSVSeq(w io.Writer, sheets []core.Sheet, exp *convert.Engine) error {
	for i, s := range sheets {
		if i > 0 {
			_, err := fmt.Fprintf(w, "--- %s ---\n", s.Name)
			if err != nil {
				return err
			}
		}
		if err := writeCSV(w, s, exp); err != nil {
			return err
		}
	}
	return nil
}

// writeCSV writes a single sheet to a writer.
// Rows are converted in source order (never via the dict map) so CSV output
// stays deterministic and duplicate keys are preserved.
func writeCSV(w io.Writer, sheet core.Sheet, exp *convert.Engine) error {
	cw := csv.NewWriter(w)

	cols := exp.CSVColumns(sheet)

	// header
	hdr := make([]string, len(cols))
	for i, c := range cols {
		hdr[i] = c.Name
	}
	if err := cw.Write(hdr); err != nil {
		return fmt.Errorf("write header: %w", err)
	}

	for ri, row := range sheet.Rows {
		m, err := exp.ConvertRow(sheet, ri, row)
		if err != nil {
			return err
		}
		rec := make([]string, len(cols))
		for i, c := range cols {
			rec[i] = formatCell(m[c.Name])
		}
		if err := cw.Write(rec); err != nil {
			return fmt.Errorf("write row: %w", err)
		}
	}

	cw.Flush()
	if err := cw.Error(); err != nil {
		return fmt.Errorf("flush csv: %w", err)
	}
	return nil
}

// formatCell converts a typed value back to its string representation for CSV output.
// Scalars use their canonical form; objects and lists are serialised as JSON.
func formatCell(v any) string {
	if v == nil {
		return ""
	}
	switch val := v.(type) {
	case bool:
		return strconv.FormatBool(val)
	case int64:
		return strconv.FormatInt(val, 10)
	case uint64:
		return strconv.FormatUint(val, 10)
	case float64:
		return strconv.FormatFloat(val, 'f', -1, 64)
	case string:
		return val
	case typing.DateValue:
		return val.Time.Format(val.Format)
	default:
		// object, list, or any other complex type → JSON
		b, err := json.Marshal(val)
		if err != nil {
			return fmt.Sprint(v)
		}
		return string(b)
	}
}
