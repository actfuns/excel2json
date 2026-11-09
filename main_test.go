package main

import (
	"io"
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestFormatName(t *testing.T) {
	if got := formatName("my_{file}_{sheet}", "data", "Sheet1"); got != "my_data_Sheet1" {
		t.Errorf("formatName = %q", got)
	}
}

func TestEncodeContent(t *testing.T) {
	for _, enc := range []string{"utf8-nobom", "other"} {
		b, err := encodeContent("hello", enc)
		if err != nil || string(b) != "hello" {
			t.Errorf("encodeContent(%q) = %q, %v", enc, b, err)
		}
	}
}

func TestFileStem(t *testing.T) {
	cases := map[string]string{
		"a.xlsx":   "a",
		"a.b.c":    "a.b",
		"noext":    "noext",
		".hidden":  ".hidden",
		"dir/x.go": "x",
	}
	for in, want := range cases {
		if got := fileStem(in); got != want {
			t.Errorf("fileStem(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestResolveTarget(t *testing.T) {
	dir := t.TempDir()
	cases := []struct {
		multi         bool
		output        string
		want          string
		wantDirectory bool
	}{
		{false, "", "", false},
		{false, filepath.Join(dir, "out.json"), filepath.Join(dir, "out.json"), false},
		{false, dir, dir, true},
		{true, filepath.Join(dir, "out.json"), filepath.Join(dir, "out.json"), true},
		{false, dir + "/", dir + "/", true},
	}
	for _, c := range cases {
		got, isDir := resolveTarget(c.multi, Options{OutputPath: c.output})
		if got != c.want || isDir != c.wantDirectory {
			t.Errorf("resolveTarget(%v, %q) = %q, %v", c.multi, c.output, got, isDir)
		}
	}
}

func TestExpandInputs(t *testing.T) {
	dir := t.TempDir()
	for _, name := range []string{"a.xlsx", "b.xls", "c.txt"} {
		if err := os.WriteFile(filepath.Join(dir, name), []byte("x"), 0644); err != nil {
			t.Fatal(err)
		}
	}
	if err := os.Mkdir(filepath.Join(dir, "sub"), 0755); err != nil {
		t.Fatal(err)
	}

	missing := filepath.Join(dir, "missing.xlsx")
	got := expandInputs([]string{dir, filepath.Join(dir, "a.xlsx"), missing})
	want := []string{filepath.Join(dir, "a.xlsx"), filepath.Join(dir, "b.xls"), missing}
	if len(got) != len(want) {
		t.Fatalf("expandInputs = %v, want %v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("expandInputs[%d] = %q, want %q", i, got[i], want[i])
		}
	}
}

func TestWriteJSONFiles(t *testing.T) {
	dir := t.TempDir()
	sheet := Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	opts := Options{ExportArray: true}

	path := filepath.Join(dir, "out.json")
	if err := writeJSON([]Sheet{sheet}, opts, path); err != nil {
		t.Fatal(err)
	}
	assertFile(t, path, `[{"id":1}]`)

	one := filepath.Join(dir, "one.json")
	if err := writeOneJSON(sheet, opts, one); err != nil {
		t.Fatal(err)
	}
	assertFile(t, one, `[{"id":1}]`)
}

func makeTwoSheetWorkbook(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	if _, err := f.NewSheet("Second"); err != nil {
		t.Fatal(err)
	}
	for _, s := range []struct {
		name string
		id   string
	}{
		{"Sheet1", "1"},
		{"Second", "2"},
	} {
		rows := [][]any{
			{"id", "v"},
			{"int", "int"},
			{s.id, "10"},
		}
		for i, r := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, i+1)
			if err := f.SetSheetRow(s.name, cell, &r); err != nil {
				t.Fatal(err)
			}
		}
	}
	path := filepath.Join(t.TempDir(), "two.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestRunOutputs(t *testing.T) {
	workbook := makeTwoSheetWorkbook(t)
	base := Options{NameRow: 0, TypeRow: 1, SkipRows: 2, ExportArray: true}

	dir := t.TempDir()
	jsonDir := filepath.Join(dir, "json") + "/"
	opts := base
	opts.Format = "json"
	opts.OutputPath = jsonDir
	opts.NameTmpl = "{sheet}"
	if err := run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "json", "Sheet1.json"), `[{"id":1,"v":10}]`)
	assertFile(t, filepath.Join(dir, "json", "Second.json"), `[{"id":2,"v":10}]`)

	merged := filepath.Join(dir, "merged.json")
	opts = base
	opts.Format = "json"
	opts.OutputPath = merged
	opts.Merge = true
	if err := run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, merged, `{"Second":[{"id":2,"v":10}],"Sheet1":[{"id":1,"v":10}]}`)

	csvDir := filepath.Join(dir, "csv") + "/"
	opts = base
	opts.Format = "csv"
	opts.OutputPath = csvDir
	if err := run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "csv", "two_Sheet1.csv"), "id,v\n1,10\n")
	assertFile(t, filepath.Join(dir, "csv", "two_Second.csv"), "id,v\n2,10\n")

	opts = base
	opts.Format = "xml"
	if err := run([]string{workbook}, opts); err == nil {
		t.Error("run with unknown format: expected error")
	}
}

func TestRunWithTypesYAMLAutoDiscovery(t *testing.T) {
	dir := t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, "types.yaml"), []byte(`
types:
  items: "$[*][1::3]:string"
`), 0644); err != nil {
		t.Fatal(err)
	}

	f := excelize.NewFile()
	rows := [][]any{
		{"id", "items"},
		{"int", "items"},
		{"1", `[["w1","i1",10],["w2","i2",20]]`},
	}
	for i, r := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := f.SetSheetRow("Sheet1", cell, &r); err != nil {
			t.Fatal(err)
		}
	}
	workbook := filepath.Join(dir, "typed.xlsx")
	if err := f.SaveAs(workbook); err != nil {
		t.Fatal(err)
	}

	opts := Options{
		NameRow:     0,
		TypeRow:     1,
		SkipRows:    2,
		ExportArray: true,
		Format:      "json",
		NameTmpl:    "{file}_{sheet}",
		OutputPath:  filepath.Join(dir, "out.json"),
	}
	if err := run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "typed_Sheet1.json"), `[{"id":1,"items":[["w1","i1",10],["w2","i2",20]]}]`)
}

func TestValidationFixture(t *testing.T) {
	fixture := filepath.Join("testdata", "validation", "validation.xlsx")
	outDir := t.TempDir()

	opts := Options{
		NameRow:     0,
		TypeRow:     1,
		SkipRows:    2,
		ExportArray: true,
		Format:      "json",
		NameTmpl:    "{sheet}",
		OutputPath:  outDir,
	}
	if err := run([]string{fixture}, opts); err != nil {
		t.Fatal(err)
	}

	assertFile(t, filepath.Join(outDir, "main.json"),
		`[{"fixed":[1,2,0,0],"id":1,"items":[["w1","i1",10],["w2","i2",20]],"kind":"ab","level":1,"packed":[5,"x",2],"refItem":"i1","score":0,"single":{"a":7,"b":"keep"}}]`)
	assertFile(t, filepath.Join(outDir, "item.json"),
		`[{"id":"i1","name":"Sword"},{"id":"i2","name":"Shield"}]`)
}

func TestWriteJSONStdout(t *testing.T) {
	old := os.Stdout
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	os.Stdout = w
	defer func() { os.Stdout = old }()

	sheet := Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	writeErr := writeJSON([]Sheet{sheet}, Options{ExportArray: true}, "")

	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	out, err := io.ReadAll(r)
	if err != nil {
		t.Fatal(err)
	}
	if writeErr != nil {
		t.Fatal(writeErr)
	}
	if want := "[{\"id\":1}]\n"; string(out) != want {
		t.Errorf("stdout = %q, want %q", out, want)
	}
}
