package app

import (
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/actfuns/excel2json/internal/core"
	"github.com/xuri/excelize/v2"
)

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

func TestEncodeContent(t *testing.T) {
	for _, enc := range []string{"", "utf8", "utf8-nobom"} {
		b, err := encodeContent("hello", enc)
		if err != nil || string(b) != "hello" {
			t.Errorf("encodeContent(%q) = %q, %v", enc, b, err)
		}
	}
	b, err := encodeContent("hello", "utf8-bom")
	if err != nil || string(b) != "\xef\xbb\xbfhello" {
		t.Errorf("encodeContent(utf8-bom) = %q, %v", b, err)
	}
	if _, err := encodeContent("hello", "gbk"); err != nil {
		t.Errorf("encodeContent(gbk): %v", err)
	}
	if _, err := encodeContent("hello", "unknown"); err == nil {
		t.Error("encodeContent(unknown): expected error")
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
		{false, filepath.Join(dir, "out"), filepath.Join(dir, "out"), true},
	}
	for _, c := range cases {
		got, isDir := resolveTarget(c.multi, core.Options{OutputPath: c.output})
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
	sheet := core.Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	opts := core.Options{ListOutput: true}

	path := filepath.Join(dir, "out.json")
	if err := writeJSON([]core.Sheet{sheet}, opts, path); err != nil {
		t.Fatal(err)
	}
	assertFile(t, path, `[{"id":1}]`)

	one := filepath.Join(dir, "one.json")
	if err := writeOneJSON(sheet, []core.Sheet{sheet}, opts, one); err != nil {
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
	base := core.Options{HeaderRow: 0, TypeRow: 1, DataRow: 2, ListOutput: true}

	dir := t.TempDir()
	jsonDir := filepath.Join(dir, "json") + "/"
	opts := base
	opts.Format = "json"
	opts.OutputPath = jsonDir
	opts.NameTmpl = "{sheet}"
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "json", "Sheet1.json"), `[{"id":1,"v":10}]`)
	assertFile(t, filepath.Join(dir, "json", "Second.json"), `[{"id":2,"v":10}]`)

	merged := filepath.Join(dir, "merged.json")
	opts = base
	opts.Format = "json"
	opts.OutputPath = merged
	opts.MergeSheets = true
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, merged, `{"Second":[{"id":2,"v":10}],"Sheet1":[{"id":1,"v":10}]}`)

	csvDir := filepath.Join(dir, "csv") + "/"
	opts = base
	opts.Format = "csv"
	opts.OutputPath = csvDir
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "csv", "two_Sheet1.csv"), "id,v\n1,10\n")
	assertFile(t, filepath.Join(dir, "csv", "two_Second.csv"), "id,v\n2,10\n")

	opts = base
	opts.Format = "xml"
	if err := Run([]string{workbook}, opts); err == nil {
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

	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{file}_{sheet}",
		OutputPath: filepath.Join(dir, "out") + "/",
	}
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "out", "typed_Sheet1.json"), `[{"id":1,"items":[["w1","i1",10],["w2","i2",20]]}]`)
}

func TestValidationFixture(t *testing.T) {
	fixture := filepath.Join("..", "..", "testdata", "validation", "validation.xlsx")
	outDir := t.TempDir()

	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{file}_{sheet}",
		OutputPath: outDir,
	}
	if err := Run([]string{fixture}, opts); err != nil {
		t.Fatal(err)
	}

	assertFile(t, filepath.Join(outDir, "validation_main.json"),
		`[{"fixed":[1,2,0,0],"id":1,"items":[["w1","i1",10],["w2","i2",20]],"kind":"ab","level":1,"packed":[5,"x",2],"refItem":"i1","score":0,"single":{"a":7,"b":"keep"}}]`)
	assertFile(t, filepath.Join(outDir, "validation_item.json"),
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

	sheet := core.Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	writeErr := writeJSON([]core.Sheet{sheet}, core.Options{ListOutput: true}, "")

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

func TestVersionFlag(t *testing.T) {
	root := NewCommand("v1.2.3")
	var buf strings.Builder
	root.SetOut(&buf)
	root.SetArgs([]string{"--version"})
	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(buf.String(), "v1.2.3") {
		t.Errorf("version output = %q", buf.String())
	}
}

func TestSingleFileTargetMergesSheets(t *testing.T) {
	workbook := makeTwoSheetWorkbook(t)
	base := core.Options{HeaderRow: 0, TypeRow: 1, DataRow: 2, ListOutput: true}
	dir := t.TempDir()

	jsonPath := filepath.Join(dir, "single.json")
	opts := base
	opts.Format = "json"
	opts.OutputPath = jsonPath
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, jsonPath, `{"Second":[{"id":2,"v":10}],"Sheet1":[{"id":1,"v":10}]}`)

	csvPath := filepath.Join(dir, "single.csv")
	opts = base
	opts.Format = "csv"
	opts.OutputPath = csvPath
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, csvPath, "id,v\n1,10\n--- Second ---\nid,v\n2,10\n")
}

func TestRunNoInputs(t *testing.T) {
	if err := Run(nil, core.Options{}); err == nil {
		t.Error("empty input list: expected error")
	}
}

func TestRunOutputDirWithoutSlash(t *testing.T) {
	workbook := makeTwoSheetWorkbook(t)
	dir := t.TempDir()
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{file}_{sheet}",
		OutputPath: filepath.Join(dir, "conf"),
	}
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "conf", "two_Sheet1.json"), `[{"id":1,"v":10}]`)
	assertFile(t, filepath.Join(dir, "conf", "two_Second.json"), `[{"id":2,"v":10}]`)
}

func TestLegacyFlags(t *testing.T) {
	workbook := filepath.Join("..", "..", "testdata", "types", "all_types.xlsx")
	dir := t.TempDir()

	root := NewCommand("v1")
	root.SetOut(io.Discard)
	root.SetErr(io.Discard)
	root.SetArgs([]string{
		"-o", filepath.Join(dir, "conf"),
		"--cell-json", "--skip-rows", "2", "--array",
		"--exclude", "#", "--pretty", "--name", "file",
		workbook,
	})
	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}
	data, err := os.ReadFile(filepath.Join(dir, "conf", "all_types.json"))
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(data), `"int": 42`) && !strings.Contains(string(data), `"int":42`) {
		t.Errorf("legacy output missing int field: %s", data)
	}
}

func TestCrossWorkbookRef(t *testing.T) {
	dir := t.TempDir()
	makeBook := func(name, sheet string, rows [][]any) string {
		f := excelize.NewFile()
		if err := f.SetSheetName("Sheet1", sheet); err != nil {
			t.Fatal(err)
		}
		for i, r := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, i+1)
			if err := f.SetSheetRow(sheet, cell, &r); err != nil {
				t.Fatal(err)
			}
		}
		path := filepath.Join(dir, name)
		if err := f.SaveAs(path); err != nil {
			t.Fatal(err)
		}
		return path
	}
	effects := makeBook("effects.xlsx", "effect", [][]any{{"id"}, {"int"}, {"12000340"}})
	buff := makeBook("buff.xlsx", "buff", [][]any{{"effects"}, {"int:ref(effect.id)"}, {"12000340"}})

	outDir := filepath.Join(dir, "out") + "/"
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{sheet}",
		OutputPath: outDir,
	}
	if err := Run([]string{buff, effects}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "out", "buff.json"), `[{"effects":12000340}]`)
}

func makeTaggedWorkbook(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	for _, s := range []struct {
		name   string
		reward int
	}{
		{"Shop@c", 10},
		{"Server@s", 20},
	} {
		if _, err := f.NewSheet(s.name); err != nil {
			t.Fatal(err)
		}
		rows := [][]any{
			{"id", "reward", "common"},
			{"int", "int", "int"},
			{"1", s.reward, "30"},
		}
		for i, r := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, i+1)
			if err := f.SetSheetRow(s.name, cell, &r); err != nil {
				t.Fatal(err)
			}
		}
	}
	path := filepath.Join(t.TempDir(), "tags.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestRunTagFiltering(t *testing.T) {
	workbook := makeTaggedWorkbook(t)
	base := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{sheet}",
	}

	dir := t.TempDir()
	opts := base
	opts.Tag = "c"
	opts.OutputPath = filepath.Join(dir, "client") + "/"
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "client", "Shop.json"), `[{"common":30,"id":1,"reward":10}]`)

	opts = base
	opts.Tag = "s"
	opts.OutputPath = filepath.Join(dir, "server") + "/"
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "server", "Server.json"), `[{"common":30,"id":1,"reward":20}]`)

	opts = base
	opts.Format = "csv"
	opts.Tag = "c"
	opts.OutputPath = filepath.Join(dir, "csv") + "/"
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "csv", "Shop.csv"), "id,reward,common\n1,10,30\n")
	if _, err := os.Stat(filepath.Join(dir, "csv", "Server.csv")); !os.IsNotExist(err) {
		t.Errorf("server sheet should not be exported to client CSV, stat err = %v", err)
	}

	opts = base
	opts.Tag = "c"
	opts.StrictTags = true
	opts.OutputPath = filepath.Join(dir, "strict") + "/"
	if err := Run([]string{workbook}, opts); err == nil {
		t.Error("strict tags with untagged columns: expected error")
	}

	opts = base
	opts.StrictTags = true
	opts.OutputPath = filepath.Join(dir, "strict-no-tag") + "/"
	if err := Run([]string{workbook}, opts); err == nil {
		t.Error("--strict-tags without --tag: expected error")
	}
}

func TestRunStrictTags(t *testing.T) {
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", "Shop@c"); err != nil {
		t.Fatal(err)
	}
	rows := [][]any{
		{"id@c", "reward@c", "common@c,s"},
		{"int", "int", "int"},
		{"1", "10", "30"},
	}
	for i, r := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := f.SetSheetRow("Shop@c", cell, &r); err != nil {
			t.Fatal(err)
		}
	}
	workbook := filepath.Join(t.TempDir(), "strict.xlsx")
	if err := f.SaveAs(workbook); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(t.TempDir(), "out") + "/"
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{sheet}",
		OutputPath: out,
		Tag:        "c",
		StrictTags: true,
	}
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(out, "Shop.json"), `[{"common":30,"id":1,"reward":10}]`)
}

func makeTaggedInput(t *testing.T, sheetName, header string) string {
	t.Helper()
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", sheetName); err != nil {
		t.Fatal(err)
	}
	rows := [][]any{
		{header},
		{"int"},
		{"1"},
	}
	for i, r := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := f.SetSheetRow(sheetName, cell, &r); err != nil {
			t.Fatal(err)
		}
	}
	path := filepath.Join(t.TempDir(), "input.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestRunTagNoMatch(t *testing.T) {
	workbook := makeTaggedInput(t, "Shop", "reward@s")
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		OutputPath: filepath.Join(t.TempDir(), "out") + "/",
		Tag:        "c",
	}
	err := Run([]string{workbook}, opts)
	if err == nil || !strings.Contains(err.Error(), `tag "c" matched no sheets or columns`) {
		t.Fatalf("Run = %v, want clear tag mismatch error", err)
	}
}

func TestRunTagSkipsUnmatchedWorkbook(t *testing.T) {
	client := makeTaggedInput(t, "Shop@c", "reward@c")
	server := makeTaggedInput(t, "Server@s", "reward@s")
	out := filepath.Join(t.TempDir(), "out") + "/"
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{sheet}",
		OutputPath: out,
		Tag:        "c",
	}
	if err := Run([]string{client, server}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(out, "Shop.json"), `[{"reward":1}]`)
	if _, err := os.Stat(filepath.Join(out, "Server.json")); !os.IsNotExist(err) {
		t.Errorf("server-only workbook should be skipped, stat err = %v", err)
	}
}

func TestRunNoTagExportsTaggedFields(t *testing.T) {
	workbook := makeTaggedWorkbook(t)
	out := filepath.Join(t.TempDir(), "out") + "/"
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{sheet}",
		OutputPath: out,
	}
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(out, "Shop.json"), `[{"common":30,"id":1,"reward":10}]`)
	assertFile(t, filepath.Join(out, "Server.json"), `[{"common":30,"id":1,"reward":20}]`)
}

func TestRunTagRefUsesFullWorkbook(t *testing.T) {
	f := excelize.NewFile()
	for _, s := range []struct {
		name   string
		header string
		typ    string
	}{
		{"Server@s", "id", "int"},
		{"Client@c", "sid", "int:ref(Server.id)"},
	} {
		if _, err := f.NewSheet(s.name); err != nil {
			t.Fatal(err)
		}
		rows := [][]any{{s.header}, {s.typ}, {"1"}}
		for i, r := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, i+1)
			if err := f.SetSheetRow(s.name, cell, &r); err != nil {
				t.Fatal(err)
			}
		}
	}
	workbook := filepath.Join(t.TempDir(), "ref-tags.xlsx")
	if err := f.SaveAs(workbook); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "out") + "/"
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{sheet}",
		OutputPath: out,
		Tag:        "c",
	}
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(out, "Client.json"), `[{"sid":1}]`)
	if _, err := os.Stat(filepath.Join(out, "Server.json")); !os.IsNotExist(err) {
		t.Errorf("server sheet should not be exported, stat err = %v", err)
	}
}

func makeSheetWorkbook(t *testing.T, sheet string, headers, types, data []string) string {
	t.Helper()
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", sheet); err != nil {
		t.Fatal(err)
	}
	rows := [][]string{headers, types, data}
	for i, row := range rows {
		vals := make([]any, len(row))
		for j, v := range row {
			vals[j] = v
		}
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := f.SetSheetRow(sheet, cell, &vals); err != nil {
			t.Fatal(err)
		}
	}
	path := filepath.Join(t.TempDir(), "input.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestRunTagKeyColumn(t *testing.T) {
	workbook := makeSheetWorkbook(t, "Shop",
		[]string{"id@s", "name@c"},
		[]string{"int", "string"},
		[]string{"1", "x"},
	)
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		Format:     "json",
		OutputPath: filepath.Join(t.TempDir(), "out") + "/",
		Tag:        "c",
		KeyColumn:  "id",
	}
	if err := Run([]string{workbook}, opts); err == nil {
		t.Error("key column tagged s under --tag c: expected error")
	}

	common := makeSheetWorkbook(t, "Shop",
		[]string{"id", "name@c"},
		[]string{"int", "string"},
		[]string{"1", "x"},
	)
	out := filepath.Join(t.TempDir(), "out") + "/"
	opts.OutputPath = out
	opts.NameTmpl = "{sheet}"
	if err := Run([]string{common}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(out, "Shop.json"), `{"1":{"id":1,"name":"x"}}`)
}

func TestRunStrictTagsCSV(t *testing.T) {
	workbook := makeSheetWorkbook(t, "Shop@c",
		[]string{"id@c", "name@c"},
		[]string{"int", "string"},
		[]string{"1", "x"},
	)
	out := filepath.Join(t.TempDir(), "out") + "/"
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		Format:     "csv",
		NameTmpl:   "{sheet}",
		OutputPath: out,
		Tag:        "c",
		StrictTags: true,
	}
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(out, "Shop.csv"), "id,name\n1,x\n")
}

func TestRunTagDuplicateSheetBase(t *testing.T) {
	f := excelize.NewFile()
	for _, s := range []struct {
		name string
		tag  string
	}{
		{"Shop@c", "c"},
		{"Shop@s", "s"},
	} {
		if _, err := f.NewSheet(s.name); err != nil {
			t.Fatal(err)
		}
		rows := [][]any{{"id@" + s.tag}, {"int"}, {"1"}}
		for i, r := range rows {
			cell, _ := excelize.CoordinatesToCellName(1, i+1)
			if err := f.SetSheetRow(s.name, cell, &r); err != nil {
				t.Fatal(err)
			}
		}
	}
	workbook := filepath.Join(t.TempDir(), "dupe.xlsx")
	if err := f.SaveAs(workbook); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "out") + "/"
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{sheet}",
		OutputPath: out,
		Tag:        "c",
	}
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(out, "Shop.json"), `[{"id":1}]`)

	opts.Tag = ""
	opts.OutputPath = filepath.Join(t.TempDir(), "out-no-tag") + "/"
	if err := Run([]string{workbook}, opts); err == nil {
		t.Error("duplicate sheet output names without --tag: expected error")
	}
}

func TestRunTagRefUsesTaggedColumn(t *testing.T) {
	workbook := makeSheetWorkbook(t, "Shop",
		[]string{"id@s", "sid@c"},
		[]string{"int", "int:ref(Shop.id)"},
		[]string{"1", "1"},
	)
	out := filepath.Join(t.TempDir(), "out") + "/"
	opts := core.Options{
		HeaderRow:  0,
		TypeRow:    1,
		DataRow:    2,
		ListOutput: true,
		Format:     "json",
		NameTmpl:   "{sheet}",
		OutputPath: out,
		Tag:        "c",
	}
	if err := Run([]string{workbook}, opts); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(out, "Shop.json"), `[{"sid":1}]`)
}
