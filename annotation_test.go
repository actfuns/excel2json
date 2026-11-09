package main

import "testing"

func TestParseFieldSpecPlainWithChecks(t *testing.T) {
	e := NewExporter(Options{})
	spec, err := e.parseFieldSpec(`int:gt(0),lt(100),ref(items.id)`)
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.rules) != 1 {
		t.Fatalf("rules = %d, want 1", len(spec.rules))
	}
	rule := spec.rules[0]
	if rule.selector != nil {
		t.Fatal("unexpected selector")
	}
	if rule.conditional {
		t.Error("rule should not be conditional")
	}
	if got := rule.typ.String(); got != "int" {
		t.Errorf("type = %q, want int", got)
	}
	if len(rule.checks) != 3 {
		t.Fatalf("checks = %d, want 3", len(rule.checks))
	}
	for i, want := range []string{"gt", "lt", "ref"} {
		if rule.checks[i].name != want {
			t.Errorf("check %d = %q, want %q", i, rule.checks[i].name, want)
		}
	}
}

func TestParseFieldSpecSelector(t *testing.T) {
	e := NewExporter(Options{})
	spec, err := e.parseFieldSpec(`$[*][1::3]:string:ref(item.id)`)
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.rules) != 1 || spec.rules[0].selector == nil {
		t.Fatal("expected one rule with selector")
	}
	rule := spec.rules[0]
	if got := rule.typ.String(); got != "string" {
		t.Errorf("type = %q, want string", got)
	}
	if len(rule.checks) != 1 || rule.checks[0].name != "ref" {
		t.Errorf("checks = %+v", rule.checks)
	}
}

func TestParseFieldSpecConditionalType(t *testing.T) {
	e := NewExporter(Options{})

	spec, err := e.parseFieldSpec(`int?:lt(1)`)
	if err != nil {
		t.Fatal(err)
	}
	if !spec.rules[0].conditional || spec.rules[0].typ.String() != "int" {
		t.Errorf("int? = %+v", spec.rules[0])
	}

	spec, err = e.parseFieldSpec(`int!:lt(1)`)
	if err != nil {
		t.Fatal(err)
	}
	if spec.rules[0].conditional || spec.rules[0].typ.String() != "int" {
		t.Errorf("int! = %+v", spec.rules[0])
	}
}

func TestParseFieldSpecMultipleRules(t *testing.T) {
	e := NewExporter(Options{})
	spec, err := e.parseFieldSpec(`list;$[0]:int32:gte(0)`)
	if err != nil {
		t.Fatal(err)
	}
	if len(spec.rules) != 2 {
		t.Fatalf("rules = %d, want 2", len(spec.rules))
	}
	if spec.rules[0].selector != nil || spec.rules[1].selector == nil {
		t.Error("unexpected selector placement")
	}
}

func TestParseFieldSpecCombinators(t *testing.T) {
	e := NewExporter(Options{})
	spec, err := e.parseFieldSpec(`int:not(gt(0),lt(10))`)
	if err != nil {
		t.Fatal(err)
	}
	ch := spec.rules[0].checks[0]
	if ch.name != "not" || len(ch.subs) != 2 {
		t.Fatalf("check = %+v", ch)
	}

	spec, err = e.parseFieldSpec(`string:or(email,phone)`)
	if err != nil {
		t.Fatal(err)
	}
	ch = spec.rules[0].checks[0]
	if ch.name != "or" || len(ch.subs) != 2 {
		t.Fatalf("check = %+v", ch)
	}
}

func TestParseFieldSpecPad(t *testing.T) {
	e := NewExporter(Options{})
	spec, err := e.parseFieldSpec(`[4]int32:pad,gte(0)`)
	if err != nil {
		t.Fatal(err)
	}
	checks := spec.rules[0].checks
	if len(checks) != 2 || checks[0].name != "pad" || checks[1].name != "gte" {
		t.Fatalf("checks = %+v", checks)
	}
}

func TestParseFieldSpecNamedType(t *testing.T) {
	e := NewExporter(Options{TypeDefs: map[string]string{
		"itemId": "int:gt(0),ref(items.id)",
	}})
	for _, annotation := range []string{"#itemId", "itemId"} {
		spec, err := e.parseFieldSpec(annotation)
		if err != nil {
			t.Errorf("%s: %v", annotation, err)
			continue
		}
		if spec.rules[0].typ.String() != "int" {
			t.Errorf("%s type = %s", annotation, spec.rules[0].typ)
		}
	}

	e = NewExporter(Options{TypeDefs: map[string]string{
		"a": "#b",
		"b": "#a",
	}})
	if _, err := e.parseFieldSpec(`#a`); err == nil {
		t.Error("cyclic reference: expected error")
	}
}

func TestParseFieldSpecErrors(t *testing.T) {
	e := NewExporter(Options{})
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
		if _, err := e.parseFieldSpec(in); err == nil {
			t.Errorf("parseFieldSpec(%q): expected error", in)
		}
	}
}
