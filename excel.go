package main

import (
	"bytes"
	"fmt"
	"os"
	"strings"

	"github.com/extrame/xls"
	"github.com/xuri/excelize/v2"
)

// rawSheet holds raw rows from a sheet before header/type processing.
type rawSheet struct {
	Name string
	Rows [][]string
}

// readExcelRows detects file format and returns raw sheet data.
// Supports both .xlsx (OOXML) and .xls (BIFF) formats.
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

	// OLE2/CFB identifier → old .xls (BIFF) format
	if bytes.HasPrefix(magic, []byte{0xd0, 0xcf, 0x11, 0xe0}) {
		return readXLS(filePath)
	}

	// Otherwise assume OOXML (ZIP-based .xlsx / .xlsm)
	return readXLSX(filePath)
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
		rows, err := f.GetRows(name, excelize.Options{RawCellValue: true})
		if err != nil {
			return nil, fmt.Errorf("read sheet [%s]: %w", name, err)
		}
		sheets = append(sheets, rawSheet{Name: name, Rows: rows})
	}

	return sheets, nil
}

// readXLS reads an old .xls (BIFF) file using extrame/xls.
func readXLS(filePath string) ([]rawSheet, error) {
	wb, err := xls.Open(filePath, "utf-8")
	if err != nil {
		return nil, fmt.Errorf("open xls file: %w", err)
	}

	var sheets []rawSheet
	for i := 0; i < wb.NumSheets(); i++ {
		ws := wb.GetSheet(i)
		if ws == nil {
			continue
		}

		rowCount := int(ws.MaxRow) + 1
		if rowCount <= 0 {
			continue
		}

		rows := make([][]string, 0, rowCount)
		for r := 0; r < rowCount; r++ {
			row, ok := safeXLSRow(ws, r)
			if !ok {
				rows = append(rows, []string{})
				continue
			}
			lastCol := row.LastCol()
			if lastCol < 0 {
				rows = append(rows, []string{})
				continue
			}
			cols := make([]string, lastCol+1)
			for c := 0; c <= lastCol; c++ {
				cols[c] = row.Col(c)
			}
			rows = append(rows, cols)
		}

		sheets = append(sheets, rawSheet{Name: ws.Name, Rows: rows})
	}

	return sheets, nil
}

// safeXLSRow wraps xls.WorkSheet.Row to handle nil pointer panics
// that occur when iterating past the last row of a sheet.
func safeXLSRow(ws *xls.WorkSheet, i int) (row *xls.Row, ok bool) {
	defer func() {
		if r := recover(); r != nil {
			ok = false
			row = nil
		}
	}()
	row = ws.Row(i)
	if row == nil {
		return nil, false
	}
	return row, true
}

// LoadExcel opens an Excel file and parses every sheet.
// Automatically detects .xlsx (OOXML) vs .xls (BIFF) format.
func LoadExcel(filePath string, opts Options) ([]Sheet, error) {
	rawSheets, err := readExcelRows(filePath)
	if err != nil {
		return nil, fmt.Errorf("open excel file: %w", err)
	}

	if len(rawSheets) == 0 {
		return nil, fmt.Errorf("excel file is empty: %s", filePath)
	}

	nameRow := opts.NameRow
	typeRow := opts.TypeRow
	header := opts.HeaderRows

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
		dataStart := header
		if dataStart <= nameRow {
			return nil, fmt.Errorf("sheet [%s]: --header=%d must be greater than --name-row=%d",
				rs.Name, header, nameRow)
		}
		if hasTypes && dataStart <= typeRow {
			return nil, fmt.Errorf("sheet [%s]: --header=%d must be greater than --type-row=%d",
				rs.Name, header, typeRow)
		}
		if dataStart < 0 {
			dataStart = 0
		}
		if dataStart > len(rows) {
			dataStart = len(rows)
		}

		dataRows := make([][]string, 0, len(rows)-dataStart)
		for ri := dataStart; ri < len(rows); ri++ {
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
