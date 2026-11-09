package export

import (
	"strings"
	"testing"

	"github.com/actfuns/excel2json/internal/core"
)

func TestJSONModes(t *testing.T) {
	s1 := core.Sheet{Name: "S", Headers: []string{"id", "v"}, Types: []string{"int", "int"}, Rows: [][]string{{"1", "10"}}}
	s2 := core.Sheet{Name: "T", Headers: []string{"id", "v"}, Types: []string{"int", "int"}, Rows: [][]string{{"2", "20"}}}

	cases := []struct {
		sheets []core.Sheet
		opts   core.Options
		want   string
	}{
		{[]core.Sheet{s1}, core.Options{ListOutput: true}, `[{"id":1,"v":10}]`},
		{[]core.Sheet{s1}, core.Options{}, `{"1":{"id":1,"v":10}}`},
		{[]core.Sheet{s1, s2}, core.Options{ListOutput: true}, `{"S":[{"id":1,"v":10}],"T":[{"id":2,"v":20}]}`},
	}
	for _, c := range cases {
		got, err := JSON(c.sheets, c.opts)
		if err != nil {
			t.Fatal(err)
		}
		if got != c.want {
			t.Errorf("JSON = %s, want %s", got, c.want)
		}
	}

	pretty, err := JSON([]core.Sheet{s1}, core.Options{ListOutput: true, Pretty: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(pretty, "\n\t") {
		t.Errorf("pretty output missing indentation: %q", pretty)
	}

	if _, err := JSON([]core.Sheet{{Name: "empty"}}, core.Options{}); err == nil {
		t.Error("JSON with no valid sheets: expected error")
	}
}
