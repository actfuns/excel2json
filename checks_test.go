package main

import (
	"encoding/json"
	"testing"
)

func parseChecksForTest(t *testing.T, e *Exporter, text string) []*checkExpr {
	t.Helper()
	checks, err := e.parseChecks(text)
	if err != nil {
		t.Fatal(err)
	}
	for _, ch := range checks {
		if err := e.compileCheck(ch, nil); err != nil {
			t.Fatal(err)
		}
	}
	return checks
}

func TestNumericChecks(t *testing.T) {
	e := NewExporter(Options{})
	pass := []struct {
		value any
		check string
	}{
		{int64(5), "gt(4)"},
		{int64(5), "gte(5)"},
		{int64(5), "lt(6)"},
		{int64(5), "lte(5)"},
		{int64(5), "between(1,5)"},
		{json.Number("5"), "eq(5)"},
		{int64(5), "positive"},
		{int64(-5), "negative"},
		{int64(5), "nonzero"},
		{int64(10), "multipleof(5)"},
	}
	for _, c := range pass {
		checks := parseChecksForTest(t, e, c.check)
		if err := e.runChecks(c.value, checks); err != nil {
			t.Errorf("%v %s: %v", c.value, c.check, err)
		}
	}

	fail := []struct {
		value any
		check string
	}{
		{int64(5), "gt(5)"},
		{int64(5), "lt(5)"},
		{int64(5), "between(6,10)"},
		{int64(5), "eq(6)"},
		{int64(5), "multipleof(2)"},
	}
	for _, c := range fail {
		checks := parseChecksForTest(t, e, c.check)
		if err := e.runChecks(c.value, checks); err == nil {
			t.Errorf("%v %s: expected failure", c.value, c.check)
		}
	}
}

func TestStringChecks(t *testing.T) {
	e := NewExporter(Options{})
	pass := []struct {
		value any
		check string
	}{
		{"中文", "len(2)"},
		{"abc", "len(2,4)"},
		{"abc", "minlen(1)"},
		{"abc", "maxlen(3)"},
		{"a@b.com", "email"},
		{"13800138000", "phone"},
		{"+8613800138000", "phone(e164)"},
		{"https://example.com/x", "url"},
		{"550e8400-e29b-41d4-a716-446655440000", "uuid"},
		{"abc123", `regex(^[a-z0-9]+$)`},
		{"pre-fix", "prefix(pre)"},
		{"pre-fix", "suffix(fix)"},
		{"pre-fix", "contains(-)"},
		{"abc", "alpha"},
		{"abc123", "alnum"},
		{"123", "numeric"},
	}
	for _, c := range pass {
		checks := parseChecksForTest(t, e, c.check)
		if err := e.runChecks(c.value, checks); err != nil {
			t.Errorf("%q %s: %v", c.value, c.check, err)
		}
	}

	fail := []struct {
		value any
		check string
	}{
		{"ab", "len(3)"},
		{"not-an-email", "email"},
		{"123", "phone"},
		{"ftp:/bad", "url"},
		{"not-a-uuid", "uuid"},
		{"ABC", `regex(^[a-z]+$)`},
		{"abc-", "suffix(x)"},
	}
	for _, c := range fail {
		checks := parseChecksForTest(t, e, c.check)
		if err := e.runChecks(c.value, checks); err == nil {
			t.Errorf("%q %s: expected failure", c.value, c.check)
		}
	}
}

func TestEnumAndCombinatorChecks(t *testing.T) {
	e := NewExporter(Options{})
	cases := []struct {
		value any
		check string
		want  bool
	}{
		{int64(2), "oneof(1,2,3)", true},
		{int64(4), "oneof(1,2,3)", false},
		{int64(4), "notoneof(1,2,3)", true},
		{"b", `oneof("a","b")`, true},
		{int64(5), "or(gt(10),lt(0))", false},
		{int64(11), "or(gt(10),lt(0))", true},
		{int64(5), "not(gt(0),lt(10))", false},
		{int64(0), "not(gt(0),lt(10))", true},
		{int64(5), "and(gt(0),lt(10))", true},
	}
	for _, c := range cases {
		checks := parseChecksForTest(t, e, c.check)
		err := e.runChecks(c.value, checks)
		if (err == nil) != c.want {
			t.Errorf("%v %s: err=%v, want pass=%v", c.value, c.check, err, c.want)
		}
	}
}

func TestArrayChecksDistribute(t *testing.T) {
	e := NewExporter(Options{})
	checks := parseChecksForTest(t, e, "gte(0)")
	if err := e.runChecks([]any{int64(1), int64(2)}, checks); err != nil {
		t.Errorf("all elements >= 0 should pass: %v", err)
	}
	if err := e.runChecks([]any{int64(1), int64(-1)}, checks); err == nil {
		t.Error("negative element should fail")
	}

	checks = parseChecksForTest(t, e, "gte(0),len(2)")
	if err := e.runChecks([]any{int64(1), int64(2)}, checks); err != nil {
		t.Errorf("element and collection checks: %v", err)
	}
	if err := e.runChecks([]any{int64(1)}, checks); err == nil {
		t.Error("collection length check should fail")
	}
}

func TestCheckTypeApplicability(t *testing.T) {
	e := NewExporter(Options{})
	checks, err := e.parseChecks("gt(0)")
	if err != nil {
		t.Fatal(err)
	}
	strType, _ := parseType("string")
	if err := e.compileChecks(checks, strType); err == nil {
		t.Error("gt should not apply to string")
	}

	checks, _ = e.parseChecks("email")
	intType, _ := parseType("int")
	if err := e.compileChecks(checks, intType); err == nil {
		t.Error("email should not apply to int")
	}
}
