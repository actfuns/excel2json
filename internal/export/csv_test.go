package export

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/actfuns/excel2json/internal/convert"
	"github.com/actfuns/excel2json/internal/core"
	"github.com/actfuns/excel2json/internal/typing"
)

func TestWriteCSVSeq(t *testing.T) {
	s1 := core.Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	s2 := core.Sheet{Name: "T", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"2"}}}

	var buf bytes.Buffer
	if err := writeCSVSeq(&buf, []core.Sheet{s1, s2}, convert.New(core.Options{})); err != nil {
		t.Fatal(err)
	}
	want := "id\n1\n--- T ---\nid\n2\n"
	if buf.String() != want {
		t.Errorf("writeCSVSeq = %q, want %q", buf.String(), want)
	}
}

func TestExportCSVFiles(t *testing.T) {
	s1 := core.Sheet{Name: "S", Headers: []string{"id", "v"}, Types: []string{"int", "string"}, Rows: [][]string{{"1", "a"}}}
	s2 := core.Sheet{Name: "T", Headers: []string{"id", "v"}, Types: []string{"int", "string"}, Rows: [][]string{{"2", "b"}}}
	opts := core.Options{}

	dir := t.TempDir()
	single := filepath.Join(dir, "single.csv")
	if _, err := CSV([]core.Sheet{s1}, opts, single); err != nil {
		t.Fatal(err)
	}
	assertFile(t, single, "id,v\n1,a\n")

	stem := filepath.Join(dir, "multi.csv")
	if _, err := CSV([]core.Sheet{s1, s2}, opts, stem); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "multi_S.csv"), "id,v\n1,a\n")
	assertFile(t, filepath.Join(dir, "multi_T.csv"), "id,v\n2,b\n")

	outDir := filepath.Join(dir, "out")
	if err := os.MkdirAll(outDir, 0755); err != nil {
		t.Fatal(err)
	}
	if _, err := CSV([]core.Sheet{s1, s2}, opts, outDir); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(outDir, "S.csv"), "id,v\n1,a\n")
	assertFile(t, filepath.Join(outDir, "T.csv"), "id,v\n2,b\n")

	if _, err := CSV(nil, opts, ""); err == nil {
		t.Error("CSV with no valid sheets: expected error")
	}
}

func assertFile(t *testing.T, path, want string) {
	t.Helper()
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if strings.ReplaceAll(string(b), "\r\n", "\n") != want {
		t.Errorf("%s = %q, want %q", path, b, want)
	}
}

func TestCSVOrderAndDuplicates(t *testing.T) {
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"id", "v"},
		Types:   []string{"int", "int"},
		Rows:    [][]string{{"1", "10"}, {"1", "20"}, {"2", "30"}},
	}
	var buf bytes.Buffer
	if err := writeCSV(&buf, sheet, convert.New(core.Options{})); err != nil {
		t.Fatal(err)
	}
	if want := "id,v\n1,10\n1,20\n2,30\n"; buf.String() != want {
		t.Errorf("CSV output = %q, want %q", buf.String(), want)
	}
}

func TestFormatCell(t *testing.T) {
	d := typing.DateValue{Time: time.Date(2024, 1, 2, 0, 0, 0, 0, time.UTC), Format: "2006/01/02"}
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

func TestCSVNameTemplate(t *testing.T) {
	s1 := core.Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	s2 := core.Sheet{Name: "T", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"2"}}}
	dir := t.TempDir()
	opts := core.Options{NameTmpl: "{file}_{sheet}", SourceStem: "data"}
	paths, err := CSV([]core.Sheet{s1, s2}, opts, filepath.Join(dir, "out.csv"))
	if err != nil {
		t.Fatal(err)
	}
	want := []string{filepath.Join(dir, "data_S.csv"), filepath.Join(dir, "data_T.csv")}
	if len(paths) != len(want) || paths[0] != want[0] || paths[1] != want[1] {
		t.Errorf("paths = %v, want %v", paths, want)
	}
	assertFile(t, want[0], "id\n1\n")
	assertFile(t, want[1], "id\n2\n")
}

func TestCSVMergeSheets(t *testing.T) {
	s1 := core.Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	s2 := core.Sheet{Name: "T", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"2"}}}
	path := filepath.Join(t.TempDir(), "merged.csv")
	if _, err := CSV([]core.Sheet{s1, s2}, core.Options{MergeSheets: true}, path); err != nil {
		t.Fatal(err)
	}
	assertFile(t, path, "id\n1\n--- T ---\nid\n2\n")
}

func TestCSVNameCollision(t *testing.T) {
	s1 := core.Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	s2 := core.Sheet{Name: "T", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"2"}}}
	if _, err := CSV([]core.Sheet{s1, s2}, core.Options{NameTmpl: "{file}", SourceStem: "data"}, filepath.Join(t.TempDir(), "out.csv")); err == nil {
		t.Error("output name collision: expected error")
	}
}
