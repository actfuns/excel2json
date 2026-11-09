package excel

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"testing"
	"time"

	"github.com/actfuns/excel2json/internal/convert"
	"github.com/actfuns/excel2json/internal/core"
	"github.com/xuri/excelize/v2"
)

func TestLoadDataRowPerSheet(t *testing.T) {
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

	opts := core.Options{HeaderRow: 0, TypeRow: -1, DataRow: 10}
	sheets, err := Load(path, opts)
	if err != nil {
		t.Fatal(err)
	}

	var second *core.Sheet
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

	opts := core.Options{HeaderRow: 0, TypeRow: 1, DataRow: 2, DateFormat: "yyyy-MM-dd", ListOutput: true}
	sheets, err := Load(path, opts)
	if err != nil {
		t.Fatal(err)
	}
	v, err := convert.New(opts).ConvertSheet(sheets[0])
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

func TestLoadTypesAndPadding(t *testing.T) {
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

	opts := core.Options{HeaderRow: 0, TypeRow: 1, DataRow: 2, ListOutput: true}
	sheets, err := Load(path, opts)
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

	v, err := convert.New(opts).ConvertSheet(sheet)
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

	sheets, err := Load(path, core.Options{HeaderRow: 0, TypeRow: 1, DataRow: 2})
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

	sheets, err := Load(path, core.Options{HeaderRow: 0, TypeRow: -1, DataRow: 1})
	if err != nil {
		t.Fatal(err)
	}
	if want := []string{"id!", "name?"}; !reflect.DeepEqual(sheets[0].Headers, want) {
		t.Errorf("Headers = %v, want %v", sheets[0].Headers, want)
	}
}

func TestLoadEmptyWorkbook(t *testing.T) {
	f := excelize.NewFile()
	if err := f.SetCellValue("Sheet1", "A1", "   "); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "blank.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if _, err := Load(path, core.Options{}); err == nil {
		t.Error("Load blank workbook: expected error")
	}
}

func convertFixture(t *testing.T, path string, opts core.Options) []map[string]any {
	t.Helper()
	sheets, err := Load(path, opts)
	if err != nil {
		t.Fatalf("%s: %v", path, err)
	}
	engine := convert.NewWithSheets(opts, sheets)
	v, err := engine.ConvertSheet(sheets[0])
	if err != nil {
		t.Fatalf("%s convert: %v", path, err)
	}
	b, err := json.Marshal(v)
	if err != nil {
		t.Fatal(err)
	}
	var rows []map[string]any
	if err := json.Unmarshal(b, &rows); err != nil {
		t.Fatal(err)
	}
	return rows
}

func TestTypeFixtures(t *testing.T) {
	opts := core.Options{HeaderRow: 0, TypeRow: 1, DataRow: 2, ListOutput: true, DateFormat: "yyyy/MM/dd"}

	all := convertFixture(t, filepath.Join("..", "..", "testdata", "types", "all_types.xlsx"), opts)
	if len(all) != 1 {
		t.Fatalf("all_types rows = %d, want 1", len(all))
	}
	if all[0]["int"] != float64(42) || all[0]["bool"] != true || all[0]["string"] != "hello" {
		t.Errorf("all_types scalar values = %v", all[0])
	}
	if all[0]["date"] != "2024/01/02" {
		t.Errorf("all_types date = %v", all[0]["date"])
	}
	if len(all[0]) != 33 {
		t.Errorf("all_types columns = %d, want 33", len(all[0]))
	}

	emptyAll := convertFixture(t, filepath.Join("..", "..", "testdata", "types", "empty_all_types.xlsx"), opts)
	if len(emptyAll) != 1 {
		t.Fatalf("empty_all_types rows = %d, want 1", len(emptyAll))
	}
	if emptyAll[0]["int"] != float64(0) || emptyAll[0]["bool"] != false || emptyAll[0]["string"] != "" {
		t.Errorf("empty_all_types scalar defaults = %v", emptyAll[0])
	}
	if emptyAll[0]["[]int"] != nil || emptyAll[0]["object"] != nil {
		t.Errorf("empty_all_types composite defaults = %v", emptyAll[0])
	}

	empty := convertFixture(t, filepath.Join("..", "..", "testdata", "types", "empty_types.xlsx"), opts)
	if len(empty) != 2 {
		t.Fatalf("empty_types rows = %d, want 2", len(empty))
	}
	if empty[0]["name"] != "alice" || empty[0]["age"] != float64(0) || empty[0]["fixed"] != nil {
		t.Errorf("empty_types row 1 = %v", empty[0])
	}
	if empty[1]["name"] != "bob" || empty[1]["age"] != float64(30) {
		t.Errorf("empty_types row 2 = %v", empty[1])
	}
}

func TestDataRowAuto(t *testing.T) {
	sheets, err := Load(filepath.Join("..", "..", "testdata", "validation", "validation.xlsx"), core.Options{
		HeaderRow: 0,
		TypeRow:   1,
		DataRow:   -1,
	})
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, sheet := range sheets {
		if sheet.Name == "main" {
			found = true
			if len(sheet.Rows) != 1 {
				t.Fatalf("auto data row: main rows = %d, want 1", len(sheet.Rows))
			}
		}
	}
	if !found {
		t.Fatal("auto data row: main sheet not found")
	}
}

func TestBlankRowsKeepRowIndices(t *testing.T) {
	f := excelize.NewFile()
	// Header at row 1, blank row 2, type row 3, data row 4.
	_ = f.SetSheetRow("Sheet1", "A1", &[]any{"id", "v"})
	_ = f.SetSheetRow("Sheet1", "A3", &[]any{"int", "int"})
	_ = f.SetSheetRow("Sheet1", "A4", &[]any{"1", "10"})

	path := filepath.Join(t.TempDir(), "gaps.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}

	sheets, err := Load(path, core.Options{HeaderRow: 0, TypeRow: 2, DataRow: -1})
	if err != nil {
		t.Fatal(err)
	}
	if len(sheets) != 1 {
		t.Fatalf("sheets = %d, want 1", len(sheets))
	}
	if got := sheets[0].Types; len(got) != 2 || got[0] != "int" || got[1] != "int" {
		t.Errorf("types = %v", got)
	}
	if len(sheets[0].Rows) != 1 || sheets[0].Rows[0][0] != "1" {
		t.Errorf("rows = %v", sheets[0].Rows)
	}
}
