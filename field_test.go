package main

import (
	"encoding/json"
	"testing"
)

func marshalValue(t *testing.T, v any) string {
	t.Helper()
	b, err := json.Marshal(v)
	if err != nil {
		t.Fatal(err)
	}
	return string(b)
}

func TestSelectorWriteBack(t *testing.T) {
	e := NewExporter(Options{})
	cell := `[[1,"a",10.0],[2,"b",20.0]]`

	v, err := e.convertCell(`$[*][2::3]:int`, cell)
	if err != nil {
		t.Fatal(err)
	}
	if got, want := marshalValue(t, v), `[[1,"a",10],[2,"b",20]]`; got != want {
		t.Errorf("write-back = %s, want %s", got, want)
	}

	v, err = e.convertCell(`$[0]:int:gte(0)`, `[5.0,"x"]`)
	if err != nil {
		t.Fatal(err)
	}
	if got, want := marshalValue(t, v), `[5,"x"]`; got != want {
		t.Errorf("single selector write-back = %s, want %s", got, want)
	}
}

func TestSingularSelectorWriteBack(t *testing.T) {
	e := NewExporter(Options{})
	v, err := e.convertCell(`$.a:int`, `{"a":5.0}`)
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != `{"a":5}` {
		t.Errorf("singular write-back = %s, want {\"a\":5}", got)
	}

	v, err = e.convertCell(`$.missing:int`, `{"a":5}`)
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != `{"a":5}` {
		t.Errorf("missing selector = %s, want {\"a\":5}", got)
	}
}

func TestConditionalTypes(t *testing.T) {
	e := NewExporter(Options{})

	pass := []struct{ cell, want string }{
		{"0", "0"},
		{"abc", `"abc"`},
	}
	annotation := `int?:lt(1);string?:minlen(1)`
	for _, c := range pass {
		v, err := e.convertCell(annotation, c.cell)
		if err != nil {
			t.Errorf("%q: %v", c.cell, err)
			continue
		}
		if got := marshalValue(t, v); got != c.want {
			t.Errorf("%q = %s, want %s", c.cell, got, c.want)
		}
	}

	// int matched, so its check failure is an error (no fallback to string).
	if _, err := e.convertCell(annotation, "5"); err == nil {
		t.Error("5: expected int check failure")
	}

	// No conditional type matches: leave the value unchanged.
	v, err := e.convertCell(`int?:lt(1);bool?:eq(true)`, "abc")
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != `"abc"` {
		t.Errorf("no-match = %s, want \"abc\"", got)
	}

	// A strict type still errors when the value does not match.
	if _, err := e.convertCell(`int!:lt(1)`, "abc"); err == nil {
		t.Error("strict int: expected error")
	}
}

func TestMultipleRulesInPlace(t *testing.T) {
	e := NewExporter(Options{})
	v, err := e.convertCell(`$[0]:int:gte(0);$[2]:int32:gte(1)`, `[5,"x",2]`)
	if err != nil {
		t.Fatal(err)
	}
	if got, want := marshalValue(t, v), `[5,"x",2]`; got != want {
		t.Errorf("in-place = %s, want %s", got, want)
	}
}

func TestWholeAndLocalRules(t *testing.T) {
	e := NewExporter(Options{})
	v, err := e.convertCell(`list;$[0]:int32:gte(0)`, `[5,"x"]`)
	if err != nil {
		t.Fatal(err)
	}
	if got, want := marshalValue(t, v), `[5,"x"]`; got != want {
		t.Errorf("list + local = %s, want %s", got, want)
	}
}

func TestPad(t *testing.T) {
	e := NewExporter(Options{})
	cases := []struct {
		typ, cell, want string
	}{
		{`[4]int32:pad,gte(0)`, `[1,2]`, `[1,2,0,0]`},
		{`[4]int32:pad`, `[1,2]`, `[1,2,0,0]`},
		{`[3]int:pad(7)`, `[1]`, `[1,7,7]`},
		{`[2]string:pad("x")`, `["a"]`, `["a","x"]`},
	}
	for _, c := range cases {
		v, err := e.convertCell(c.typ, c.cell)
		if err != nil {
			t.Errorf("%s %s: %v", c.typ, c.cell, err)
			continue
		}
		if got := marshalValue(t, v); got != c.want {
			t.Errorf("%s %s = %s, want %s", c.typ, c.cell, got, c.want)
		}
	}

	if _, err := e.convertCell(`[2]int:pad`, `[1,2,3]`); err == nil {
		t.Error("too many elements: expected error")
	}
	if _, err := e.convertCell(`[2]int:pad,gte(0)`, `[1,-1]`); err == nil {
		t.Error("element check: expected error")
	}
}

func TestHeaderRequiredOverride(t *testing.T) {
	e := NewExporter(Options{})
	required := true
	nullable := false

	if _, err := e.convertCellMeta("int", "", &required); err == nil {
		t.Error("required empty cell: expected error")
	}
	v, err := e.convertCellMeta("int", "", &nullable)
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != "0" {
		t.Errorf("nullable empty = %s, want 0", got)
	}

	if _, err := e.convertCellMeta("int:gt(0)", "", &required); err == nil {
		t.Error("required empty checked cell: expected error")
	}
	v, err = e.convertCellMeta("int:gt(0)", "", &nullable)
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != "0" {
		t.Errorf("nullable empty checked = %s, want 0", got)
	}
}

func TestSplitFieldMarker(t *testing.T) {
	cases := []struct {
		in       string
		want     string
		nullable *bool
	}{
		{"id", "id", nil},
		{"id!", "id", boolPtr(true)},
		{"id?", "id", boolPtr(false)},
	}
	for _, c := range cases {
		got, req := splitFieldMarker(c.in)
		if got != c.want {
			t.Errorf("splitFieldMarker(%q) = %q, want %q", c.in, got, c.want)
		}
		if (c.nullable == nil) != (req == nil) || (c.nullable != nil && *c.nullable != *req) {
			t.Errorf("splitFieldMarker(%q) marker = %v, want %v", c.in, req, c.nullable)
		}
	}
}

func boolPtr(v bool) *bool { return &v }

func TestNamedTypeReference(t *testing.T) {
	e := NewExporter(Options{TypeDefs: map[string]string{
		"itemId": "int:gte(0)",
	}})
	v, err := e.convertCell(`#itemId`, `5`)
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != "5" {
		t.Errorf("named type = %s, want 5", got)
	}
}

func TestTypeMarkerDoesNotAffectEmpty(t *testing.T) {
	e := NewExporter(Options{})
	v, err := e.convertCell(`int!`, "")
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != "0" {
		t.Errorf("int! empty = %s, want 0", got)
	}

	required := true
	if _, err := e.convertCellMeta(`int!`, "", &required); err == nil {
		t.Error("header required: expected error")
	}

	e = NewExporter(Options{Required: true})
	if _, err := e.convertCell(`int!`, ""); err == nil {
		t.Error("--required: expected error")
	}
}

func TestRequiredOptionAppliesToSpec(t *testing.T) {
	e := NewExporter(Options{Required: true})
	if _, err := e.convertCell(`int:gt(0)`, ""); err == nil {
		t.Error("--required empty checked cell: expected error")
	}

	// Header-level nullable overrides --required.
	nullable := false
	v, err := e.convertCellMeta(`int:gt(0)`, "", &nullable)
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != "0" {
		t.Errorf("nullable checked cell = %s, want 0", got)
	}
}

func TestOverlappingRulesAllowSameType(t *testing.T) {
	e := NewExporter(Options{})
	v, err := e.convertCell(`$[0]:int:gte(0);$[0]:int:lte(10)`, `[5]`)
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != `[5]` {
		t.Errorf("overlapping rules = %s, want [5]", got)
	}
	if _, err := e.convertCell(`$[0]:int:gte(0);$[0]:int:lte(10)`, `[11]`); err == nil {
		t.Error("second overlapping rule should fail")
	}
}

func TestRefCheckElementWise(t *testing.T) {
	sheets := []Sheet{
		{
			Name:    "item",
			Headers: []string{"id"},
			Types:   []string{"int"},
			Rows:    [][]string{{"1"}, {"3"}},
		},
	}
	e := NewExporterWithSheets(Options{}, sheets)
	if _, err := e.convertCell(`$[1::2]:int:ref(item.id)`, `[0,1,0,3]`); err != nil {
		t.Fatalf("element-wise ref: %v", err)
	}
	if _, err := e.convertCell(`$[1::2]:int:ref(item.id)`, `[0,1,0,2]`); err == nil {
		t.Error("missing element ref: expected error")
	}
}

func TestRefUsesOutputPackageName(t *testing.T) {
	sheets := []Sheet{
		{Name: "item", Headers: []string{"id"}, Types: []string{"string"}, Rows: [][]string{{"i1"}}},
		{Name: "main", Headers: []string{"item"}, Types: []string{"string:ref(data_item.id)"}, Rows: [][]string{{"i1"}}},
	}
	e := NewExporterWithSheets(Options{NameTmpl: "{file}_{sheet}", SourceName: "data"}, sheets)
	if _, err := e.ConvertRow(sheets[1], 0, sheets[1].Rows[0]); err != nil {
		t.Fatalf("package ref: %v", err)
	}

	// The .json suffix and the raw sheet name are also accepted.
	for _, ref := range []string{"data_item.json", "item"} {
		e := NewExporterWithSheets(Options{NameTmpl: "{file}_{sheet}", SourceName: "data"}, []Sheet{
			{Name: "item", Headers: []string{"id"}, Types: []string{"string"}, Rows: [][]string{{"i1"}}},
			{Name: "main", Headers: []string{"item"}, Types: []string{"string:ref(" + ref + ".id)"}, Rows: [][]string{{"i1"}}},
		})
		if _, err := e.ConvertRow(e.sheets[1], 0, e.sheets[1].Rows[0]); err != nil {
			t.Errorf("ref(%s.id): %v", ref, err)
		}
	}

	// A template without {sheet} makes packages ambiguous.
	ambiguous := NewExporterWithSheets(Options{NameTmpl: "{file}", SourceName: "data"}, []Sheet{
		{Name: "a"},
		{Name: "b"},
	})
	if _, err := ambiguous.findRefSheet("data"); err == nil {
		t.Error("ambiguous package: expected error")
	}
}

func TestRefCheck(t *testing.T) {
	sheets := []Sheet{
		{
			Name:    "item",
			Headers: []string{"id"},
			Types:   []string{"int"},
			Rows:    [][]string{{"1"}, {"2"}},
		},
		{
			Name:    "main",
			Headers: []string{"id", "itemId"},
			Types:   []string{"int", "int:ref(item.id)"},
			Rows:    [][]string{{"1", "2"}, {"2", "3"}},
		},
	}
	e := NewExporterWithSheets(Options{}, sheets)

	row, err := e.ConvertRow(sheets[1], 0, sheets[1].Rows[0])
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, row["itemId"]); got != "2" {
		t.Errorf("itemId = %s, want 2", got)
	}

	if _, err := e.ConvertRow(sheets[1], 1, sheets[1].Rows[1]); err == nil {
		t.Error("missing foreign key: expected error")
	}

	noContext := NewExporter(Options{})
	if _, err := noContext.ConvertRow(sheets[1], 0, sheets[1].Rows[0]); err == nil {
		t.Error("ref without workbook context: expected error")
	}
}
