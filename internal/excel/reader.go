package excel

import (
	"bytes"
	"fmt"
	"os"
	"strings"

	"github.com/actfuns/excel2json/internal/core"
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
		sheets = append(sheets, rawSheet{Name: name, Rows: rawRows})
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
			row := make([]string, len(vals))
			copy(row, vals)
			rows = append(rows, row)
		}
		if err := s.Err(); err != nil {
			return nil, fmt.Errorf("sheet [%s]: %w", name, err)
		}

		sheets = append(sheets, rawSheet{Name: name, Rows: rows})
	}

	return sheets, nil
}

// isEmptyRow reports whether a row has no non-blank cells.
func isEmptyRow(vals []string) bool {
	for _, v := range vals {
		if strings.TrimSpace(v) != "" {
			return false
		}
	}
	return true
}

// Load opens an Excel file and parses every sheet.
func Load(filePath string, opts core.Options) ([]core.Sheet, error) {
	rawSheets, err := readExcelRows(filePath)
	if err != nil {
		return nil, err
	}

	if len(rawSheets) == 0 {
		return nil, fmt.Errorf("excel file is empty: %s", filePath)
	}

	headerRow := opts.HeaderRow
	typeRow := opts.TypeRow
	dataRow := opts.DataRow

	var sheets []core.Sheet

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

		if typeRow >= 0 && typeRow <= headerRow {
			return nil, fmt.Errorf("sheet [%s]: --type-row=%d must be greater than --header-row=%d",
				rs.Name, typeRow, headerRow)
		}
		hasTypes := typeRow >= 0 && typeRow < len(rows) && typeRow > headerRow

		// --- column names from nameRow ---
		headers := make([]string, maxCols)
		required := make([]*bool, maxCols)
		for i := range headers {
			headers[i] = fmt.Sprintf("col_%d", i)
		}
		if headerRow >= 0 && headerRow < len(rows) {
			for ci, val := range rows[headerRow] {
				if v := strings.TrimSpace(val); v != "" && ci < len(headers) {
					if hasTypes {
						name, req := core.SplitFieldMarker(v)
						headers[ci] = name
						required[ci] = req
					} else {
						headers[ci] = v
					}
				}
			}
		}

		// --- type annotations from typeRow ---
		types := make([]string, maxCols)
		if hasTypes {
			for ci, val := range rows[typeRow] {
				if v := strings.TrimSpace(val); v != "" && ci < len(types) {
					types[ci] = v
				}
			}
		}

		// --- data rows ---
		if dataRow < 0 {
			dataRow = headerRow
			if hasTypes && typeRow > dataRow {
				dataRow = typeRow
			}
			dataRow++
		}
		if dataRow <= headerRow {
			return nil, fmt.Errorf("sheet [%s]: --data-row=%d must be greater than --header-row=%d",
				rs.Name, dataRow, headerRow)
		}
		if hasTypes && dataRow <= typeRow {
			return nil, fmt.Errorf("sheet [%s]: --data-row=%d must be greater than --type-row=%d",
				rs.Name, dataRow, typeRow)
		}
		sheetSkip := dataRow
		if sheetSkip < 0 {
			sheetSkip = 0
		}
		if sheetSkip > len(rows) {
			sheetSkip = len(rows)
		}

		dataRows := make([][]string, 0, len(rows)-sheetSkip)
		for ri := sheetSkip; ri < len(rows); ri++ {
			row := rows[ri]
			if isEmptyRow(row) {
				continue
			}
			padded := make([]string, maxCols)
			copy(padded, row)
			dataRows = append(dataRows, padded)
		}
		if len(dataRows) == 0 {
			continue
		}

		sheets = append(sheets, core.Sheet{
			Name:     rs.Name,
			Headers:  headers,
			Types:    types,
			Required: required,
			Rows:     dataRows,
		})
	}

	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in: %s", filePath)
	}

	return sheets, nil
}
