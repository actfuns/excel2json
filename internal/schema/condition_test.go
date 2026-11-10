package schema

import "testing"

func TestCondOpString(t *testing.T) {
	cases := map[CondOp]string{
		OpEq: "=", OpNe: "!=", OpGt: ">", OpGte: ">=", OpLt: "<", OpLte: "<=",
		CondOp(99): "?",
	}
	for op, want := range cases {
		if got := op.String(); got != want {
			t.Errorf("%d.String() = %q, want %q", op, got, want)
		}
	}
}

func TestCompareValue(t *testing.T) {
	cases := []struct {
		value any
		op    CondOp
		lit   string
		want  bool
	}{
		{int64(5), OpEq, "5", true},
		{int64(5), OpNe, "4", true},
		{int64(5), OpGt, "4", true},
		{int64(5), OpGte, "5", true},
		{int64(5), OpLt, "6", true},
		{int64(5), OpLte, "5", true},
		{"b", OpGt, "a", true},
		{"a", OpEq, "'a'", true},
		{true, OpEq, "true", true},
		{nil, OpEq, "null", true},
		{nil, OpNe, "null", false},
	}
	for _, c := range cases {
		got, err := CompareValue(c.value, c.op, c.lit)
		if err != nil {
			t.Errorf("CompareValue(%v, %v, %q): %v", c.value, c.op, c.lit, err)
			continue
		}
		if got != c.want {
			t.Errorf("CompareValue(%v, %v, %q) = %v, want %v", c.value, c.op, c.lit, got, c.want)
		}
	}

	if _, err := CompareValue(int64(1), OpGt, "not-a-number"); err == nil {
		t.Error("invalid numeric literal: expected error")
	}
	if _, err := CompareValue([]any{1}, OpEq, "1"); err == nil {
		t.Error("array equality: expected error")
	}
}

func TestIsPlainAndZeroForSpec(t *testing.T) {
	plain, err := Parse("int", nil)
	if err != nil {
		t.Fatal(err)
	}
	if !IsPlain(plain) {
		t.Error("int should be plain")
	}
	if got := ZeroForSpec(plain); got != int64(0) {
		t.Errorf("ZeroForSpec(int) = %#v, want int64(0)", got)
	}

	for _, annotation := range []string{"int:gt(0)", "$[0]:int", "when(type=1) int", "int;string"} {
		spec, err := Parse(annotation, nil)
		if err != nil {
			t.Fatalf("%s: %v", annotation, err)
		}
		if IsPlain(spec) {
			t.Errorf("%s should not be plain", annotation)
		}
	}

	checkOnly, err := Parse("gt(0)", nil)
	if err != nil {
		t.Fatal(err)
	}
	if got := ZeroForSpec(checkOnly); got != "" {
		t.Errorf("ZeroForSpec(check-only) = %#v, want empty string", got)
	}

	padded, err := Parse("[3]int:pad", nil)
	if err != nil {
		t.Fatal(err)
	}
	if !HasPad(padded.Rules[0].Checks) {
		t.Error("pad check not detected")
	}
	arr, ok := ZeroForSpec(padded).([]any)
	if !ok || len(arr) != 3 {
		t.Errorf("ZeroForSpec([3]int:pad) = %#v, want []any len 3", ZeroForSpec(padded))
	}
}
