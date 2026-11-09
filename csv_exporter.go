package main

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// exportCSV writes all valid sheets from the given sheets slice.
//   - No output path → stdout.
//   - Single sheet → one file.
//   - Multiple sheets → auto-directory: if the path looks like a file path,
//     it's used as a stem and multiple files are created alongside it.
func exportCSV(sheets []Sheet, opts Options, outPath string) error {
	exp := NewExporterWithSheets(opts, sheets)
	valid := exp.FilterSheets(sheets)
	if len(valid) == 0 {
		return fmt.Errorf("no valid sheets to export")
	}

	if outPath == "" {
		return writeCSVSeq(os.Stdout, valid, exp)
	}

	// single sheet → single file
	if len(valid) == 1 {
		return writeCSVFile(outPath, valid[0], exp)
	}

	// multiple sheets → auto directory mode
	if !isDirPath(outPath) {
		// treat outPath as a stem: /tmp/out/multi.csv → /tmp/out/multi_Sheet1.csv
		dir := filepath.Dir(outPath)
		stem := strings.TrimSuffix(filepath.Base(outPath), ".csv")
		if err := os.MkdirAll(dir, 0755); err != nil {
			return fmt.Errorf("mkdir %s: %w", dir, err)
		}
		for _, s := range valid {
			name := sanitizeName(s.Name)
			p := filepath.Join(dir, stem+"_"+name+".csv")
			if err := writeCSVFile(p, s, exp); err != nil {
				return fmt.Errorf("write sheet [%s] csv: %w", s.Name, err)
			}
		}
		return nil
	}

	// directory path
	if err := os.MkdirAll(outPath, 0755); err != nil {
		return fmt.Errorf("mkdir: %w", err)
	}
	for _, s := range valid {
		name := sanitizeName(s.Name)
		p := filepath.Join(outPath, name+".csv")
		if err := writeCSVFile(p, s, exp); err != nil {
			return fmt.Errorf("write sheet [%s] csv: %w", s.Name, err)
		}
	}
	return nil
}

// writeCSVFile writes a single sheet to a CSV file.
func writeCSVFile(path string, sheet Sheet, exp *Exporter) error {
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("create csv: %w", err)
	}
	defer f.Close()
	return writeCSV(f, sheet, exp)
}

// writeCSVSeq writes all sheets sequentially to a writer.
// Multiple sheets are separated by comment lines.
func writeCSVSeq(w io.Writer, sheets []Sheet, exp *Exporter) error {
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
func writeCSV(w io.Writer, sheet Sheet, exp *Exporter) error {
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
	case dateValue:
		return val.t.Format(val.format)
	default:
		// object, list, or any other complex type → JSON
		b, err := json.Marshal(val)
		if err != nil {
			return fmt.Sprint(v)
		}
		return string(b)
	}
}

func sanitizeName(name string) string {
	r := strings.NewReplacer(
		"/", "_", "\\", "_", ":", "_", "*", "_",
		"?", "_", "\"", "_", "<", "_", ">", "_", "|", "_", " ", "_",
	)
	return r.Replace(name)
}

// isDirPath returns true when the path looks like a directory.
func isDirPath(path string) bool {
	if strings.HasSuffix(path, "/") || strings.HasSuffix(path, "\\") {
		return true
	}
	fi, err := os.Stat(path)
	return err == nil && fi.IsDir()
}
