package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// exportCSV writes all valid sheets from the given sheets slice.
//   - No output path → stdout.
//   - Single sheet → one file.
//   - Multiple sheets → auto-directory: if the path looks like a file path,
//     it's used as a stem and multiple files are created alongside it.
func exportCSV(sheets []Sheet, opts Options, outPath string) error {
	exp := NewExporter(opts)
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
		os.MkdirAll(dir, 0755)
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
func writeCSV(w io.Writer, sheet Sheet, exp *Exporter) error {
	cw := csv.NewWriter(w)
	defer cw.Flush()

	cols := exp.CSVColumns(sheet)

	// header
	hdr := make([]string, len(cols))
	for i, c := range cols {
		hdr[i] = c.Name
	}
	if err := cw.Write(hdr); err != nil {
		return fmt.Errorf("write header: %w", err)
	}

	// data rows
	for _, row := range sheet.Rows {
		rec := make([]string, len(cols))
		for i, c := range cols {
			cell := ""
			if c.Index < len(row) {
				cell = strings.TrimSpace(row[c.Index])
			}
			if cell == "" {
				rec[i] = fmt.Sprintf("%v", exp.columnDefault(sheet, c.Index))
			} else {
				rec[i] = cell
			}
		}
		if err := cw.Write(rec); err != nil {
			return fmt.Errorf("write row: %w", err)
		}
	}
	return nil
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
