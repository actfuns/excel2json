package main

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestCSVOrderAndDuplicates(t *testing.T) {
	sheet := Sheet{
		Name:    "S",
		Headers: []string{"id", "v"},
		Types:   []string{"int", "int"},
		Rows: [][]string{
			{"1", "10"},
			{"1", "20"},
			{"2", "30"},
		},
	}

	var buf bytes.Buffer
	if err := writeCSV(&buf, sheet, NewExporter(Options{})); err != nil {
		t.Fatal(err)
	}

	want := "id,v\n1,10\n1,20\n2,30\n"
	if buf.String() != want {
		t.Errorf("CSV output = %q, want %q", buf.String(), want)
	}
}

func TestCSVEmptyHeaderMatchesJSON(t *testing.T) {
	sheet := Sheet{
		Name:    "S",
		Headers: []string{"skip_me", ""},
		Types:   []string{"int", "int"},
		Rows:    [][]string{{"1", "2"}},
	}
	exp := NewExporter(Options{ExcludePrefix: []string{"skip_"}})

	cols := exp.CSVColumns(sheet)
	if len(cols) != 1 || cols[0].Name != "col_1" {
		t.Fatalf("CSVColumns = %+v, want one column named col_1", cols)
	}

	row, err := exp.ConvertRow(sheet, 0, sheet.Rows[0])
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := row["col_1"]; !ok {
		t.Errorf("ConvertRow keys = %v, want col_1", row)
	}
}

func TestLegacyConvert(t *testing.T) {
	cases := []struct {
		opts       Options
		cell, want string
	}{
		{Options{}, "12.0", "12"},
		{Options{}, "12.5", "12.5"},
		{Options{}, "abc", `"abc"`},
		{Options{}, "NaN", `"NaN"`},
		{Options{}, "+Inf", `"+Inf"`},
		{Options{}, "15699.999999999998", "15700"},
		{Options{}, `{"a":1}`, `"{\"a\":1}"`},
		{Options{CellJSON: true}, `{"a":1}`, `{"a":1}`},
		{Options{CellJSON: true, AllString: true}, `{"a":1}`, `"{\"a\":1}"`},
		{Options{AllString: true}, "12.0", `"12"`},
		{Options{AllString: true}, "12.5", `"12.5"`},
	}
	for _, c := range cases {
		v, err := NewExporter(c.opts).legacyConvert(c.cell)
		if err != nil {
			t.Errorf("%q: %v", c.cell, err)
			continue
		}
		b, _ := json.Marshal(v)
		if string(b) != c.want {
			t.Errorf("legacyConvert(%q) = %s, want %s", c.cell, b, c.want)
		}
	}
}

func TestFilterSheets(t *testing.T) {
	exp := NewExporter(Options{ExcludePrefix: []string{"skip_"}})
	sheets := []Sheet{
		{Name: "ok", Headers: []string{"id"}, Rows: [][]string{{"1"}}},
		{Name: "skip_me", Headers: []string{"id"}, Rows: [][]string{{"1"}}},
		{Name: "empty", Headers: []string{"id"}},
		{Name: "noheaders", Rows: [][]string{{"1"}}},
	}
	got := exp.FilterSheets(sheets)
	if len(got) != 1 || got[0].Name != "ok" {
		t.Errorf("FilterSheets = %v, want only ok", got)
	}
}

func TestConvertSheetModes(t *testing.T) {
	sheet := Sheet{
		Name:    "S",
		Headers: []string{"id", "v"},
		Types:   []string{"int", "int"},
		Rows:    [][]string{{"1", "10"}, {"2", "20"}},
	}
	cases := []struct {
		opts Options
		want string
	}{
		{Options{ExportArray: true}, `[{"id":1,"v":10},{"id":2,"v":20}]`},
		{Options{}, `{"1":{"id":1,"v":10},"2":{"id":2,"v":20}}`},
		{Options{KeyColumn: "v"}, `{"10":{"id":1,"v":10},"20":{"id":2,"v":20}}`},
	}
	for _, c := range cases {
		v, err := NewExporter(c.opts).ConvertSheet(sheet)
		if err != nil {
			t.Fatal(err)
		}
		b, _ := json.Marshal(v)
		if string(b) != c.want {
			t.Errorf("ConvertSheet = %s, want %s", b, c.want)
		}
	}
}

func TestFormatCell(t *testing.T) {
	d := dateValue{t: time.Date(2024, 1, 2, 0, 0, 0, 0, time.UTC), format: "2006/01/02"}
	cases := []struct {
		v    any
		want string
	}{
		{nil, ""},
		{true, "true"},
		{int64(-5), "-5"},
		{uint64(18446744073709551615), "18446744073709551615"},
		{float64(1.5), "1.5"},
		{"x", "x"},
		{d, "2024/01/02"},
		{[]any{int64(1), "a"}, `[1,"a"]`},
		{map[string]any{"a": int64(1)}, `{"a":1}`},
	}
	for _, c := range cases {
		if got := formatCell(c.v); got != c.want {
			t.Errorf("formatCell(%v) = %q, want %q", c.v, got, c.want)
		}
	}
}

func TestSanitizeNameAndDirPath(t *testing.T) {
	if got, want := sanitizeName(`a/b\c:d*e?f"g<h>i|j k`), "a_b_c_d_e_f_g_h_i_j_k"; got != want {
		t.Errorf("sanitizeName = %q, want %q", got, want)
	}

	dir := t.TempDir()
	if !isDirPath(dir) || !isDirPath(dir+"/") {
		t.Error("isDirPath should be true for a directory and a trailing slash")
	}
	file := filepath.Join(dir, "x.json")
	if err := os.WriteFile(file, []byte("{}"), 0644); err != nil {
		t.Fatal(err)
	}
	if isDirPath(file) {
		t.Error("isDirPath should be false for a file")
	}
	if isDirPath(filepath.Join(dir, "missing")) {
		t.Error("isDirPath should be false for a missing path")
	}
}
