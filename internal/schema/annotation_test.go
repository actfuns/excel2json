package schema

import "testing"

func TestParseAnnotationPlainWithChecks(t *testing.T) {
	spec, err := Parse(`int:gt(0),lt(100),ref(items.id)`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.Rules) != 1 {
		t.Fatalf("rules = %d, want 1", len(spec.Rules))
	}
	rule := spec.Rules[0]
	if rule.Selector != nil {
		t.Fatal("unexpected selector")
	}
	if got := rule.Type.String(); got != "int" {
		t.Errorf("type = %q, want int", got)
	}
	if len(rule.Checks) != 3 {
		t.Fatalf("checks = %d, want 3", len(rule.Checks))
	}
	for i, want := range []string{"gt", "lt", "ref"} {
		if rule.Checks[i].Name != want {
			t.Errorf("check %d = %q, want %q", i, rule.Checks[i].Name, want)
		}
	}
}

func TestParseAnnotationSelector(t *testing.T) {
	spec, err := Parse(`$[*][1::3]:string:ref(item.id)`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.Rules) != 1 || spec.Rules[0].Selector == nil {
		t.Fatal("expected one rule with selector")
	}
	rule := spec.Rules[0]
	if got := rule.Type.String(); got != "string" {
		t.Errorf("type = %q, want string", got)
	}
	if len(rule.Checks) != 1 || rule.Checks[0].Name != "ref" {
		t.Errorf("checks = %+v", rule.Checks)
	}
}

func TestParseAnnotationGuards(t *testing.T) {
	spec, err := Parse(`when(type=1, level>=10, typeof($.kind) = [3]int) []int`, nil)
	if err != nil {
		t.Fatal(err)
	}
	rule := spec.Rules[0]
	if len(rule.Guards) != 3 {
		t.Fatalf("guards = %d, want 3", len(rule.Guards))
	}
	if rule.Guards[0].Ref.Column != "type" || rule.Guards[0].Op != OpEq || rule.Guards[0].Literal != "1" {
		t.Errorf("guard 0 = %+v", rule.Guards[0])
	}
	if rule.Guards[1].Op != OpGte {
		t.Errorf("guard 1 = %+v", rule.Guards[1])
	}
	if rule.Guards[2].Kind != CondType || rule.Guards[2].TypeExpr.String() != "[3]int" {
		t.Errorf("guard 2 = %+v", rule.Guards[2])
	}

	spec, err = Parse(`when(kind.a=2) int:gt(0)`, nil)
	if err != nil {
		t.Fatal(err)
	}
	g := spec.Rules[0].Guards[0]
	if g.Ref.Column != "kind" || g.Ref.Path == nil {
		t.Errorf("column path guard = %+v", g)
	}

	spec, err = Parse(`when(typeof($) != null) int`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if spec.Rules[0].Guards[0].Kind != CondType || spec.Rules[0].Guards[0].Op != OpNe {
		t.Errorf("typeof ne guard = %+v", spec.Rules[0].Guards[0])
	}

	spec, err = Parse(`when(type($) = int) int`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if spec.Rules[0].Guards[0].Kind != CondType {
		t.Errorf("type() alias guard = %+v", spec.Rules[0].Guards[0])
	}

	spec, err = Parse(`when(or(type=1, level>=10)) int`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if spec.Rules[0].Guards[0].Kind != CondOr || len(spec.Rules[0].Guards[0].Subs) != 2 {
		t.Errorf("or guard = %+v", spec.Rules[0].Guards[0])
	}

	spec, err = Parse(`when(not(type=1)) int`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if spec.Rules[0].Guards[0].Kind != CondNot {
		t.Errorf("not guard = %+v", spec.Rules[0].Guards[0])
	}
}

func TestParseAnnotationMultipleRules(t *testing.T) {
	spec, err := Parse(`list;$[0]:int32:gte(0)`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.Rules) != 2 {
		t.Fatalf("rules = %d, want 2", len(spec.Rules))
	}
	if spec.Rules[0].Selector != nil || spec.Rules[1].Selector == nil {
		t.Error("unexpected selector placement")
	}
}

func TestParseAnnotationCombinators(t *testing.T) {
	spec, err := Parse(`int:not(gt(0),lt(10))`, nil)
	if err != nil {
		t.Fatal(err)
	}
	ch := spec.Rules[0].Checks[0]
	if ch.Name != "not" || len(ch.Subs) != 2 {
		t.Fatalf("check = %+v", ch)
	}
}

func TestParseAnnotationPad(t *testing.T) {
	spec, err := Parse(`[4]int32:pad,gte(0)`, nil)
	if err != nil {
		t.Fatal(err)
	}
	checks := spec.Rules[0].Checks
	if len(checks) != 2 || checks[0].Name != "pad" || checks[1].Name != "gte" {
		t.Fatalf("checks = %+v", checks)
	}
}

func TestParseAnnotationNamedType(t *testing.T) {
	defs := map[string]string{"itemId": "int:gt(0),ref(items.id)"}
	for _, annotation := range []string{"#itemId", "itemId"} {
		spec, err := Parse(annotation, defs)
		if err != nil {
			t.Errorf("%s: %v", annotation, err)
			continue
		}
		if spec.Rules[0].Type.String() != "int" {
			t.Errorf("%s type = %s", annotation, spec.Rules[0].Type)
		}
	}

	if _, err := Parse(`#a`, map[string]string{"a": "#b", "b": "#a"}); err == nil {
		t.Error("cyclic reference: expected error")
	}
}

func TestParseAnnotationErrors(t *testing.T) {
	bad := []string{
		"",
		";",
		"string:gt(0)",
		"int:pad",
		"[4]int32:pad(1,2)",
		"int:unknown(1)",
		"int:not()",
		"$[0]",
		"#missing",
		"int:gt(0) || string:len(2)",
		"$[0]:int?:gt(0)",
		"when(type) int",
		"when(type=) int",
		"when(typeof($)) int",
		"when(type=1)",
	}
	for _, in := range bad {
		if _, err := Parse(in, nil); err == nil {
			t.Errorf("Parse(%q): expected error", in)
		}
	}
}

func TestParseAnnotationConditionOperators(t *testing.T) {
	cases := []struct {
		in  string
		op  CondOp
		lit string
	}{
		{"when(type=1) int", OpEq, "1"},
		{"when(type==1) int", OpEq, "1"},
		{"when(type!=0) int", OpNe, "0"},
		{"when(level>=10) int", OpGte, "10"},
		{"when(level<=10) int", OpLte, "10"},
		{"when(level>1) int", OpGt, "1"},
		{"when(level<1) int", OpLt, "1"},
	}
	for _, c := range cases {
		spec, err := Parse(c.in, nil)
		if err != nil {
			t.Errorf("%s: %v", c.in, err)
			continue
		}
		g := spec.Rules[0].Guards[0]
		if g.Op != c.op || g.Literal != c.lit {
			t.Errorf("%s = %v %q, want %v %q", c.in, g.Op, g.Literal, c.op, c.lit)
		}
	}
}
