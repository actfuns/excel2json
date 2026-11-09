package convert

import (
	"encoding/json"
	"testing"

	"github.com/actfuns/excel2json/internal/core"
)

func TestCSVEmptyHeaderMatchesJSON(t *testing.T) {
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"skip_me", ""},
		Types:   []string{"int", "int"},
		Rows:    [][]string{{"1", "2"}},
	}
	exp := New(core.Options{ExcludePrefixes: []string{"skip_"}})

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

func TestFilterSheets(t *testing.T) {
	exp := New(core.Options{ExcludePrefixes: []string{"skip_"}})
	sheets := []core.Sheet{
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
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"id", "v"},
		Types:   []string{"int", "int"},
		Rows:    [][]string{{"1", "10"}, {"2", "20"}},
	}
	cases := []struct {
		opts core.Options
		want string
	}{
		{core.Options{ListOutput: true}, `[{"id":1,"v":10},{"id":2,"v":20}]`},
		{core.Options{}, `{"1":{"id":1,"v":10},"2":{"id":2,"v":20}}`},
		{core.Options{KeyColumn: "v"}, `{"10":{"id":1,"v":10},"20":{"id":2,"v":20}}`},
	}
	for _, c := range cases {
		v, err := New(c.opts).ConvertSheet(sheet)
		if err != nil {
			t.Fatal(err)
		}
		b, _ := json.Marshal(v)
		if string(b) != c.want {
			t.Errorf("ConvertSheet = %s, want %s", b, c.want)
		}
	}
}

func TestDictDuplicateKey(t *testing.T) {
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"id", "v"},
		Types:   []string{"int", "int"},
		Rows:    [][]string{{"1", "10"}, {"1", "20"}},
	}
	// Default: later row replaces the earlier one.
	v, err := New(core.Options{}).ConvertSheet(sheet)
	if err != nil {
		t.Fatal(err)
	}
	row := v.(map[string]any)["1"].(map[string]any)
	if row["v"] != int64(20) {
		t.Errorf("last-wins value = %v, want 20", row["v"])
	}
	// Strict mode: duplicates are an error.
	if _, err := New(core.Options{UniqueKeys: true}).ConvertSheet(sheet); err == nil {
		t.Error("unique-keys: expected error")
	}
}

func TestDictMissingKeyColumn(t *testing.T) {
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"id", "v"},
		Types:   []string{"int", "int"},
		Rows:    [][]string{{"1", "10"}},
	}
	if _, err := New(core.Options{KeyColumn: "missing"}).ConvertSheet(sheet); err == nil {
		t.Error("missing key column: expected error")
	}
}

func TestDictKeyColumnCaseInsensitive(t *testing.T) {
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"id", "v"},
		Types:   []string{"int", "int"},
		Rows:    [][]string{{"1", "10"}},
	}
	v, err := New(core.Options{KeyColumn: "ID"}).ConvertSheet(sheet)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := v.(map[string]any)["1"]; !ok {
		t.Errorf("keyed output = %v", v)
	}
}

func TestDateDefaultFormat(t *testing.T) {
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"d"},
		Types:   []string{"date"},
		Rows:    [][]string{{"45293"}},
	}
	v, err := New(core.Options{ListOutput: true}).ConvertSheet(sheet)
	if err != nil {
		t.Fatal(err)
	}
	b, _ := json.Marshal(v)
	if got, want := string(b), `[{"d":"2024/01/02"}]`; got != want {
		t.Errorf("date default = %s, want %s", got, want)
	}
}
