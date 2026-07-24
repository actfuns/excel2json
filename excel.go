package main

import (
	"bytes"
	"fmt"
	"os"
	"strings"

	"github.com/pbnjay/grate"
	_ "github.com/pbnjay/grate/xls"
	"github.com/xuri/excelize/v2"
)

// rawSheet holds raw rows from a sheet before header/type processing.
type rawSheet struct {
	Name string
	Rows [][]string
}

// readExcelRows detects file format and returns raw sheet data.
// .xlsx → excelize (for true OOXML),  .xls → grate (for BIFF/WPS).
func readExcelRows(filePath string) ([]rawSheet, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}

	magic := make([]byte, 8)
	if _, err := f.Read(magic); err != nil {
		f.Close()
		return nil, fmt.Errorf("read file header: %w", err)
	}
	f.Close()

	// True OOXML (.xlsx) — use excelize
	if bytes.HasPrefix(magic, []byte{'P', 'K', 0x03, 0x04}) {
		return readXLSX(filePath)
	}

	// OLE2/CFB (.xls / WPS) — use grate
	if bytes.HasPrefix(magic, []byte{0xd0, 0xcf, 0x11, 0xe0}) {
		return readXLS(filePath)
	}

	return nil, fmt.Errorf("%s: unsupported file format", filePath)
}

// readXLSX reads a true .xlsx file using excelize.
func readXLSX(filePath string) ([]rawSheet, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	names := f.GetSheetList()
	var sheets []rawSheet

	for _, name := range names {
		rawRows, err := f.GetRows(name, excelize.Options{RawCellValue: true})
		if err != nil {
			return nil, fmt.Errorf("read sheet [%s]: %w", name, err)
		}
		rows := make([][]string, 0, len(rawRows))
		for _, r := range rawRows {
			if isEmptyRow(r) {
				continue
			}
			rows = append(rows, r)
		}
		sheets = append(sheets, rawSheet{Name: name, Rows: rows})
	}

	return sheets, nil
}

// readXLS reads an old .xls (BIFF) file using grate.
func readXLS(filePath string) ([]rawSheet, error) {
	wb, err := grate.Open(filePath)
	if err != nil {
		return nil, fmt.Errorf("open xls file: %w", err)
	}
	defer wb.Close()

	names, err := wb.List()
	if err != nil {
		return nil, err
	}

	var sheets []rawSheet
	for _, name := range names {
		s, err := wb.Get(name)
		if err != nil {
			return nil, fmt.Errorf("sheet [%s]: %w", name, err)
		}

		var rows [][]string
		for s.Next() {
			vals := s.Strings()
			if isEmptyRow(vals) {
				continue
			}
			row := make([]string, len(vals))
			copy(row, vals)
			rows = append(rows, row)
		}

		sheets = append(sheets, rawSheet{Name: name, Rows: rows})
	}

	return sheets, nil
}

// isEmptyRow reports whether a row has no non-empty cells.
func isEmptyRow(vals []string) bool {
	for _, v := range vals {
		if v != "" {
			return false
		}
	}
	return true
}

// LoadExcel opens an Excel file and parses every sheet.
func LoadExcel(filePath string, opts Options) ([]Sheet, error) {
	rawSheets, err := readExcelRows(filePath)
	if err != nil {
		return nil, err
	}

	if len(rawSheets) == 0 {
		return nil, fmt.Errorf("excel file is empty: %s", filePath)
	}

	nameRow := opts.NameRow
	typeRow := opts.TypeRow
	skipRows := opts.SkipRows

	var sheets []Sheet

	for _, rs := range rawSheets {
		rows := rs.Rows
		if len(rows) == 0 {
			continue
		}

		// determine max column count across all rows
		maxCols := 0
		for _, row := range rows {
			if len(row) > maxCols {
				maxCols = len(row)
			}
		}
		if maxCols == 0 {
			continue
		}

		// --- column names from nameRow ---
		headers := make([]string, maxCols)
		for i := range headers {
			headers[i] = fmt.Sprintf("col_%d", i)
		}
		if nameRow >= 0 && nameRow < len(rows) {
			for ci, val := range rows[nameRow] {
				if v := strings.TrimSpace(val); v != "" && ci < len(headers) {
					headers[ci] = v
				}
			}
		}

		// --- type annotations from typeRow ---
		types := make([]string, maxCols)
		hasTypes := typeRow >= 0 && typeRow < len(rows) && typeRow > nameRow
		if hasTypes {
			for ci, val := range rows[typeRow] {
				if v := strings.TrimSpace(val); v != "" && ci < len(types) {
					types[ci] = strings.ToLower(v)
				}
			}
		}

		// --- data rows ---
		if skipRows <= nameRow {
			return nil, fmt.Errorf("sheet [%s]: --skip=%d must be greater than --name-row=%d",
				rs.Name, skipRows, nameRow)
		}
		if hasTypes && skipRows <= typeRow {
			return nil, fmt.Errorf("sheet [%s]: --skip=%d must be greater than --type-row=%d",
				rs.Name, skipRows, typeRow)
		}
		if skipRows < 0 {
			skipRows = 0
		}
		if skipRows > len(rows) {
			skipRows = len(rows)
		}

		dataRows := make([][]string, 0, len(rows)-skipRows)
		for ri := skipRows; ri < len(rows); ri++ {
			row := rows[ri]
			padded := make([]string, maxCols)
			copy(padded, row)
			dataRows = append(dataRows, padded)
		}

		sheets = append(sheets, Sheet{
			Name:    rs.Name,
			Headers: headers,
			Types:   types,
			Rows:    dataRows,
		})
	}

	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in: %s", filePath)
	}

	return sheets, nil
}
