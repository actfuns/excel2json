package main

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// LoadExcel opens an Excel file and parses every sheet.
// nameRow specifies which row contains column names (0-based).
// typeRow specifies which row contains type annotations (-1 to disable).
// dataStart = nameRow + headerSkip.
func LoadExcel(filePath string, opts Options) ([]Sheet, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("open excel file: %w", err)
	}
	defer f.Close()

	names := f.GetSheetList()
	if len(names) == 0 {
		return nil, fmt.Errorf("excel file is empty: %s", filePath)
	}

	nameRow := opts.NameRow
	typeRow := opts.TypeRow
	header := opts.HeaderRows

	var sheets []Sheet

	for _, sheetName := range names {
		rows, err := f.GetRows(sheetName, excelize.Options{RawCellValue: true})
		if err != nil {
			return nil, fmt.Errorf("read sheet [%s]: %w", sheetName, err)
		}
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
				sheetName, header, nameRow)
		}
		if hasTypes && dataStart <= typeRow {
			return nil, fmt.Errorf("sheet [%s]: --header=%d must be greater than --type-row=%d",
				sheetName, header, typeRow)
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
			Name:    sheetName,
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
