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
	if rule.Conditional {
		t.Error("rule should not be conditional")
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

func TestParseAnnotationConditionalType(t *testing.T) {
	spec, err := Parse(`int?:lt(1)`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !spec.Rules[0].Conditional || spec.Rules[0].Type.String() != "int" {
		t.Errorf("int? = %+v", spec.Rules[0])
	}

	spec, err = Parse(`int!:lt(1)`, nil)
	if err != nil {
		t.Fatal(err)
	}
	if spec.Rules[0].Conditional || spec.Rules[0].Type.String() != "int" {
		t.Errorf("int! = %+v", spec.Rules[0])
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
	}
	for _, in := range bad {
		if _, err := Parse(in, nil); err == nil {
			t.Errorf("Parse(%q): expected error", in)
		}
	}
}
