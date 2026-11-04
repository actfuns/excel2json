package main

import (
	"strings"
	"testing"
)

func TestExportJSONModes(t *testing.T) {
	s1 := Sheet{Name: "S", Headers: []string{"id", "v"}, Types: []string{"int", "int"}, Rows: [][]string{{"1", "10"}}}
	s2 := Sheet{Name: "T", Headers: []string{"id", "v"}, Types: []string{"int", "int"}, Rows: [][]string{{"2", "20"}}}

	cases := []struct {
		sheets []Sheet
		opts   Options
		want   string
	}{
		{[]Sheet{s1}, Options{ExportArray: true}, `[{"id":1,"v":10}]`},
		{[]Sheet{s1}, Options{}, `{"1":{"id":1,"v":10}}`},
		{[]Sheet{s1, s2}, Options{ExportArray: true}, `{"S":[{"id":1,"v":10}],"T":[{"id":2,"v":20}]}`},
	}
	for _, c := range cases {
		got, err := ExportJSON(c.sheets, c.opts)
		if err != nil {
			t.Fatal(err)
		}
		if got != c.want {
			t.Errorf("ExportJSON = %s, want %s", got, c.want)
		}
	}

	pretty, err := ExportJSON([]Sheet{s1}, Options{ExportArray: true, Pretty: true})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(pretty, "\n\t") {
		t.Errorf("pretty output missing indentation: %q", pretty)
	}

	if _, err := ExportJSON([]Sheet{{Name: "empty"}}, Options{}); err == nil {
		t.Error("ExportJSON with no valid sheets: expected error")
	}
}
