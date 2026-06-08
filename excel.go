package main

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// LoadExcel opens an Excel file and parses every sheet.
// The first headerRow rows of each sheet are treated as column names;
// subsequent rows become data rows.
func LoadExcel(filePath string, headerRow int) ([]Sheet, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("open excel file: %w", err)
	}
	defer f.Close()

	names := f.GetSheetList()
	if len(names) == 0 {
		return nil, fmt.Errorf("excel file is empty: %s", filePath)
	}

	var sheets []Sheet

	for _, name := range names {
		rows, err := f.GetRows(name)
		if err != nil {
			return nil, fmt.Errorf("read sheet [%s]: %w", name, err)
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

		// build column names from the first row only — matches C# ExcelDataReader UseHeaderRow=true
		headers := make([]string, maxCols)
		for i := range headers {
			headers[i] = fmt.Sprintf("col_%d", i)
		}
		if len(rows) > 0 {
			for ci, val := range rows[0] {
				if v := strings.TrimSpace(val); v != "" && ci < len(headers) {
					headers[ci] = v
				}
			}
		}

		// data rows start from headerRow (0-indexed).
		// Rows 1..headerRow-1 are skipped (extra header info, e.g. Chinese descriptions).
		dataStart := headerRow
		if dataStart < 1 {
			dataStart = 1
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
			Name:    name,
			Headers: headers,
			Rows:    dataRows,
		})
	}

	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in: %s", filePath)
	}

	return sheets, nil
}
