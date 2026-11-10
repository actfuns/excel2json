package convert

import (
	"encoding/json"
	"testing"

	"github.com/actfuns/excel2json/internal/core"
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
	e := New(core.Options{})
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
	e := New(core.Options{})
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

func TestGuards(t *testing.T) {
	e := New(core.Options{ListOutput: true})
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"type", "level", "kind", "a", "b", "c", "d", "e", "f", "g", "h", "k"},
		Types: []string{
			"int", "int", "object",
			"when(type=1) int:gt(0)",
			"when(type=1, level>=10) string:len(1)",
			"when(kind.a=2) int:gt(0)",
			"when(typeof($) = [3]int) [3]int",
			"when(typeof($) = string) string:len(2,4)",
			"when(type!=0) int:gte(0)",
			"when(typeof(kind.b) = string) string:len(1)",
			"when(typeof(kind.arr) = [3]int) [3]int",
			"when($.flag=true) $.v:int",
		},
	}
	row := []string{"1", "10", `{"a":2,"b":"k","arr":[1,2,3]}`, "5", "x", "7", "[1,2,3]", "abc", "8", "z", "[4,5,6]", `{"flag":true,"v":5}`}
	m, err := e.ConvertRow(sheet, 0, row)
	if err != nil {
		t.Fatal(err)
	}
	cases := map[string]string{
		"a": "5",
		"b": `"x"`,
		"c": "7",
		"d": "[1,2,3]",
		"e": `"abc"`,
		"f": "8",
		"g": `"z"`,
		"h": "[4,5,6]",
		"k": `{"flag":true,"v":5}`,
	}
	for key, want := range cases {
		if got := marshalValue(t, m[key]); got != want {
			t.Errorf("%s = %s, want %s", key, got, want)
		}
	}

	// Skipped guards leave the raw cell untouched.
	row2 := []string{"2", "5", `{"a":3}`, "bad", "y", "bad", "[1]", "abc", "9", "z", "[4]", `{"flag":false,"v":5}`}
	m, err = e.ConvertRow(sheet, 1, row2)
	if err != nil {
		t.Fatal(err)
	}
	for key, want := range map[string]string{
		"a": `"bad"`,
		"b": `"y"`,
		"c": `"bad"`,
		"d": "[1]",
	} {
		if got := marshalValue(t, m[key]); got != want {
			t.Errorf("skipped %s = %s, want %s", key, got, want)
		}
	}
}

func TestGuardErrors(t *testing.T) {
	e := New(core.Options{})
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"type", "a"},
		Types:   []string{"int", "when(missing=1) int"},
	}
	if _, err := e.ConvertRow(sheet, 0, []string{"1", "5"}); err == nil {
		t.Error("missing guard column: expected error")
	}
}

func TestMultipleRulesInPlace(t *testing.T) {
	e := New(core.Options{})
	v, err := e.convertCell(`$[0]:int:gte(0);$[2]:int32:gte(1)`, `[5,"x",2]`)
	if err != nil {
		t.Fatal(err)
	}
	if got, want := marshalValue(t, v), `[5,"x",2]`; got != want {
		t.Errorf("in-place = %s, want %s", got, want)
	}
}

func TestWholeAndLocalRules(t *testing.T) {
	e := New(core.Options{})
	v, err := e.convertCell(`list;$[0]:int32:gte(0)`, `[5,"x"]`)
	if err != nil {
		t.Fatal(err)
	}
	if got, want := marshalValue(t, v), `[5,"x"]`; got != want {
		t.Errorf("list + local = %s, want %s", got, want)
	}
}

func TestPad(t *testing.T) {
	e := New(core.Options{})
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
	e := New(core.Options{})
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

func TestNamedTypeReference(t *testing.T) {
	e := New(core.Options{TypeDefs: map[string]string{
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

func TestEmptyType(t *testing.T) {
	e := New(core.Options{})
	v, err := e.convertCell(`int`, "")
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, v); got != "0" {
		t.Errorf("int empty = %s, want 0", got)
	}

	required := true
	if _, err := e.convertCellMeta(`int`, "", &required); err == nil {
		t.Error("header required: expected error")
	}

	e = New(core.Options{RequireByDefault: true})
	if _, err := e.convertCell(`int`, ""); err == nil {
		t.Error("--required-by-default: expected error")
	}
}

func TestRequiredOptionAppliesToSpec(t *testing.T) {
	e := New(core.Options{RequireByDefault: true})
	if _, err := e.convertCell(`int:gt(0)`, ""); err == nil {
		t.Error("--required-by-default empty checked cell: expected error")
	}

	// Header-level nullable overrides --required-by-default.
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
	e := New(core.Options{})
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
	sheets := []core.Sheet{
		{
			Name:    "item",
			Headers: []string{"id"},
			Types:   []string{"int"},
			Rows:    [][]string{{"1"}, {"3"}},
		},
	}
	e := NewWithSheets(core.Options{}, sheets)
	if _, err := e.convertCell(`$[1::2]:int:ref(item.id)`, `[0,1,0,3]`); err != nil {
		t.Fatalf("element-wise ref: %v", err)
	}
	if _, err := e.convertCell(`$[1::2]:int:ref(item.id)`, `[0,1,0,2]`); err == nil {
		t.Error("missing element ref: expected error")
	}
}

func TestRefUsesOutputPackageName(t *testing.T) {
	sheets := []core.Sheet{
		{Name: "item", Headers: []string{"id"}, Types: []string{"string"}, Rows: [][]string{{"i1"}}},
		{Name: "main", Headers: []string{"item"}, Types: []string{"string:ref(data_item.id)"}, Rows: [][]string{{"i1"}}},
	}
	e := NewWithSheets(core.Options{NameTmpl: "{file}_{sheet}", SourceStem: "data"}, sheets)
	if _, err := e.ConvertRow(sheets[1], 0, sheets[1].Rows[0]); err != nil {
		t.Fatalf("package ref: %v", err)
	}

	// The .json suffix is accepted; the raw sheet name is not when a template
	// produces a different output name.
	for _, ref := range []string{"data_item.json"} {
		e := NewWithSheets(core.Options{NameTmpl: "{file}_{sheet}", SourceStem: "data"}, []core.Sheet{
			{Name: "item", Headers: []string{"id"}, Types: []string{"string"}, Rows: [][]string{{"i1"}}},
			{Name: "main", Headers: []string{"item"}, Types: []string{"string:ref(" + ref + ".id)"}, Rows: [][]string{{"i1"}}},
		})
		if _, err := e.ConvertRow(e.sheets[1].Sheet, 0, e.sheets[1].Sheet.Rows[0]); err != nil {
			t.Errorf("ref(%s.id): %v", ref, err)
		}
	}

	// The raw sheet name is rejected when a template is configured.
	strict := NewWithSheets(core.Options{NameTmpl: "{file}_{sheet}", SourceStem: "data"}, []core.Sheet{
		{Name: "item", Headers: []string{"id"}, Types: []string{"string"}, Rows: [][]string{{"i1"}}},
		{Name: "main", Headers: []string{"item"}, Types: []string{"string:ref(item.id)"}, Rows: [][]string{{"i1"}}},
	})
	if _, err := strict.ConvertRow(strict.sheets[1].Sheet, 0, strict.sheets[1].Sheet.Rows[0]); err == nil {
		t.Error("raw sheet name should not match a templated output name")
	}

	// A template without {sheet} makes packages ambiguous.
	ambiguous := NewWithSheets(core.Options{NameTmpl: "{file}", SourceStem: "data"}, []core.Sheet{
		{Name: "a"},
		{Name: "b"},
	})
	if _, err := ambiguous.findRefSheet("data"); err == nil {
		t.Error("ambiguous package: expected error")
	}
}

func TestRefCheck(t *testing.T) {
	sheets := []core.Sheet{
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
	e := NewWithSheets(core.Options{}, sheets)

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

	noContext := New(core.Options{})
	if _, err := noContext.ConvertRow(sheets[1], 0, sheets[1].Rows[0]); err == nil {
		t.Error("ref without workbook context: expected error")
	}
}

func TestCrossWorkbookRef(t *testing.T) {
	effect := core.Sheet{Name: "effect", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"12000340"}}}
	buff := core.Sheet{Name: "buff", Headers: []string{"effects"}, Types: []string{"int:ref(effect.id)"}, Rows: [][]string{{"12000340"}, {"1"}}}
	e := NewWithContext(core.Options{}, []SheetContext{
		{Sheet: effect, OutputName: "effect"},
		{Sheet: buff, OutputName: "buff"},
	})
	if _, err := e.ConvertRow(buff, 0, buff.Rows[0]); err != nil {
		t.Fatalf("cross-workbook ref: %v", err)
	}
	if _, err := e.ConvertRow(buff, 1, buff.Rows[1]); err == nil {
		t.Error("missing cross-workbook ref: expected error")
	}
}

func TestGuardCombinators(t *testing.T) {
	e := New(core.Options{})
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"type", "a", "b"},
		Types: []string{
			"int",
			"when(or(type=1, type=2)) int:gt(0)",
			"when(not(type=1)) int:gt(0)",
		},
	}
	cases := []struct {
		row   []string
		wantA string
		wantB string
	}{
		{[]string{"2", "5", "6"}, "5", "6"},
		{[]string{"1", "5", "6"}, "5", `"6"`},
		{[]string{"3", "bad", "7"}, `"bad"`, "7"},
	}
	for _, c := range cases {
		m, err := e.ConvertRow(sheet, 0, c.row)
		if err != nil {
			t.Errorf("%v: %v", c.row, err)
			continue
		}
		if got := marshalValue(t, m["a"]); got != c.wantA {
			t.Errorf("%v a = %s, want %s", c.row, got, c.wantA)
		}
		if got := marshalValue(t, m["b"]); got != c.wantB {
			t.Errorf("%v b = %s, want %s", c.row, got, c.wantB)
		}
	}
}

func TestGuardExcelBoolean(t *testing.T) {
	e := New(core.Options{})
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"active", "a", "b"},
		Types: []string{
			"bool",
			"when(active=true) int:gt(0)",
			"when(typeof($) = bool) bool",
		},
	}
	m, err := e.ConvertRow(sheet, 0, []string{"TRUE", "5", "FALSE"})
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, m["a"]); got != "5" {
		t.Errorf("a = %s, want 5", got)
	}
	if got := marshalValue(t, m["b"]); got != "false" {
		t.Errorf("b = %s, want false", got)
	}
}

func TestGuardNotShortCircuits(t *testing.T) {
	e := New(core.Options{})
	sheet := core.Sheet{
		Name:    "S",
		Headers: []string{"type", "a"},
		Types:   []string{"int", "when(not(type=1, missing=2)) int"},
	}
	m, err := e.ConvertRow(sheet, 0, []string{"0", "5"})
	if err != nil {
		t.Fatal(err)
	}
	if got := marshalValue(t, m["a"]); got != "5" {
		t.Errorf("a = %s, want 5", got)
	}
}
