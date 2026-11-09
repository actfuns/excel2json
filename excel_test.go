package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

func TestLoadExcelSkipRowsPerSheet(t *testing.T) {
	f := excelize.NewFile()
	_, err := f.NewSheet("Second")
	if err != nil {
		t.Fatal(err)
	}

	write := func(sheet string, rows [][]any) {
		for i, r := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, i+1)
			if err := f.SetSheetRow(sheet, cell, &r); err != nil {
				t.Fatal(err)
			}
		}
	}

	write("Sheet1", [][]any{{"id"}, {"1"}})
	rows := [][]any{{"id"}}
	for i := 0; i <= 10; i++ {
		rows = append(rows, []any{fmt.Sprintf("%d", i)})
	}
	write("Second", rows)

	path := filepath.Join(t.TempDir(), "multi.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	opts := Options{NameRow: 0, TypeRow: -1, SkipRows: 10}
	sheets, err := LoadExcel(path, opts)
	if err != nil {
		t.Fatal(err)
	}

	var second *Sheet
	for i := range sheets {
		if sheets[i].Name == "Second" {
			second = &sheets[i]
		}
	}
	if second == nil {
		t.Fatal("Second sheet not found")
	}
	if len(second.Rows) != 2 {
		t.Fatalf("Second rows = %d, want 2: %v", len(second.Rows), second.Rows)
	}
	if second.Rows[0][0] != "9" || second.Rows[1][0] != "10" {
		t.Errorf("Second rows = %v, want [[9] [10]]", second.Rows)
	}
}

func TestXLSXDateSerial(t *testing.T) {
	f := excelize.NewFile()
	_ = f.SetSheetRow("Sheet1", "A1", &[]any{"id", "d"})
	_ = f.SetSheetRow("Sheet1", "A2", &[]any{"int", "date"})
	_ = f.SetSheetRow("Sheet1", "A3", &[]any{"1", "d"})
	_ = f.SetSheetRow("Sheet1", "A4", &[]any{"2", "d"})

	dates := []time.Time{
		time.Date(2024, 1, 2, 0, 0, 0, 0, time.UTC),
		time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC),
	}
	style, err := f.NewStyle(&excelize.Style{NumFmt: 14})
	if err != nil {
		t.Fatal(err)
	}
	for i, d := range dates {
		cell := fmt.Sprintf("B%d", i+3)
		if err := f.SetCellValue("Sheet1", cell, d); err != nil {
			t.Fatal(err)
		}
		if err := f.SetCellStyle("Sheet1", cell, cell, style); err != nil {
			t.Fatal(err)
		}
	}

	path := filepath.Join(t.TempDir(), "dates.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	opts := Options{NameRow: 0, TypeRow: 1, SkipRows: 2, DateFormat: "yyyy-MM-dd", ExportArray: true}
	sheets, err := LoadExcel(path, opts)
	if err != nil {
		t.Fatal(err)
	}
	v, err := NewExporter(opts).ConvertSheet(sheets[0])
	if err != nil {
		t.Fatal(err)
	}
	b, _ := json.Marshal(v)
	want := `[{"d":"2024-01-02","id":1},{"d":"1900-01-01","id":2}]`
	if string(b) != want {
		t.Errorf("date serial output = %s, want %s", b, want)
	}
}

func TestReadExcelRowsErrors(t *testing.T) {
	if _, err := readExcelRows(filepath.Join(t.TempDir(), "missing.xlsx")); err == nil {
		t.Error("readExcelRows missing file: expected error")
	}

	path := filepath.Join(t.TempDir(), "bad.xlsx")
	if err := os.WriteFile(path, []byte("not-an-excel-file"), 0644); err != nil {
		t.Fatal(err)
	}
	if _, err := readExcelRows(path); err == nil {
		t.Error("readExcelRows unsupported format: expected error")
	}
}

func TestIsEmptyRow(t *testing.T) {
	if !isEmptyRow(nil) || !isEmptyRow([]string{"", "  \t"}) {
		t.Error("blank rows should be empty")
	}
	if isEmptyRow([]string{"", "x"}) {
		t.Error("non-blank row should not be empty")
	}
}

func TestLoadExcelTypesAndPadding(t *testing.T) {
	f := excelize.NewFile()
	rows := [][]any{
		{"ID", "Value", "Extra"},
		{"INT", "[]STRING", "UINT"},
		{"1", `["a"]`},
	}
	for i, r := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := f.SetSheetRow("Sheet1", cell, &r); err != nil {
			t.Fatal(err)
		}
	}
	path := filepath.Join(t.TempDir(), "types.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	opts := Options{NameRow: 0, TypeRow: 1, SkipRows: 2, ExportArray: true}
	sheets, err := LoadExcel(path, opts)
	if err != nil {
		t.Fatal(err)
	}
	sheet := sheets[0]
	if len(sheet.Rows) != 1 || len(sheet.Rows[0]) != 3 || sheet.Rows[0][2] != "" {
		t.Fatalf("row padding = %v", sheet.Rows)
	}
	wantTypes := []string{"INT", "[]STRING", "UINT"}
	for i, want := range wantTypes {
		if sheet.Types[i] != want {
			t.Errorf("Types[%d] = %q, want %q", i, sheet.Types[i], want)
		}
	}

	v, err := NewExporter(opts).ConvertSheet(sheet)
	if err != nil {
		t.Fatal(err)
	}
	b, _ := json.Marshal(v)
	if want := `[{"Extra":0,"ID":1,"Value":["a"]}]`; string(b) != want {
		t.Errorf("converted = %s, want %s", b, want)
	}
}

func TestHeaderRequiredMarkers(t *testing.T) {
	f := excelize.NewFile()
	rows := [][]any{
		{"id!", "name?", "plain"},
		{"int", "string", "int"},
		{"1", "x", "2"},
	}
	for i, r := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := f.SetSheetRow("Sheet1", cell, &r); err != nil {
			t.Fatal(err)
		}
	}
	path := filepath.Join(t.TempDir(), "markers.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	sheets, err := LoadExcel(path, Options{NameRow: 0, TypeRow: 1, SkipRows: 2})
	if err != nil {
		t.Fatal(err)
	}
	sheet := sheets[0]
	if want := []string{"id", "name", "plain"}; !reflect.DeepEqual(sheet.Headers, want) {
		t.Errorf("Headers = %v, want %v", sheet.Headers, want)
	}
	if len(sheet.Required) != 3 {
		t.Fatalf("Required length = %d", len(sheet.Required))
	}
	if sheet.Required[0] == nil || !*sheet.Required[0] {
		t.Errorf("id! should be required")
	}
	if sheet.Required[1] == nil || *sheet.Required[1] {
		t.Errorf("name? should be nullable")
	}
	if sheet.Required[2] != nil {
		t.Errorf("plain should not have a marker")
	}
}

func TestHeaderMarkersIgnoredWithoutTypeRow(t *testing.T) {
	f := excelize.NewFile()
	rows := [][]any{
		{"id!", "name?"},
		{"1", "x"},
	}
	for i, r := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := f.SetSheetRow("Sheet1", cell, &r); err != nil {
			t.Fatal(err)
		}
	}
	path := filepath.Join(t.TempDir(), "untyped_markers.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	sheets, err := LoadExcel(path, Options{NameRow: 0, TypeRow: -1, SkipRows: 1})
	if err != nil {
		t.Fatal(err)
	}
	if want := []string{"id!", "name?"}; !reflect.DeepEqual(sheets[0].Headers, want) {
		t.Errorf("Headers = %v, want %v", sheets[0].Headers, want)
	}
}

func TestLoadExcelEmptyWorkbook(t *testing.T) {
	f := excelize.NewFile()
	if err := f.SetCellValue("Sheet1", "A1", "   "); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "blank.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if _, err := LoadExcel(path, Options{}); err == nil {
		t.Error("LoadExcel blank workbook: expected error")
	}
}
