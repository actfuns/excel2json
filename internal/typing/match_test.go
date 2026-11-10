package typing

import (
	"encoding/json"
	"math"
	"testing"
)

func TestMatchJSON(t *testing.T) {
	cases := []struct {
		typ  string
		val  any
		want bool
	}{
		{"int", json.Number("1"), true},
		{"int", json.Number("1.5"), false},
		{"int8", json.Number("127"), true},
		{"int8", json.Number("128"), false},
		{"int64", json.Number("9223372036854775807"), true},
		{"uint64", json.Number("18446744073709551615"), true},
		{"uint64", json.Number("-1"), false},
		{"float", json.Number("1.5"), true},
		{"string", "abc", true},
		{"bool", true, true},
		{"null", nil, true},
		{"null", "x", false},
		{"[3]int", []any{json.Number("1"), json.Number("2"), json.Number("3")}, true},
		{"[3]int", []any{json.Number("1"), json.Number("2")}, false},
		{"[]string", []any{"a", "b"}, true},
		{"[]string", []any{"a", json.Number("1")}, false},
		{"map[string]int", map[string]any{"a": json.Number("1")}, true},
		{"map[string]int", map[string]any{"a": "x"}, false},
		{"[2][2]int", []any{[]any{json.Number("1")}, []any{json.Number("2")}}, false},
		{"any", map[string]any{}, true},
		{"int", int64(1), true},
		{"int8", int64(128), false},
		{"uint", uint(1), true},
		{"uint", int8(1), true},
		{"uint8", uint64(256), false},
		{"uint8", float32(2), true},
		{"float32", float64(1.5), true},
		{"float32", float64(math.MaxFloat32) * 2, false},
		{"map[int]string", map[string]any{"1": "a"}, true},
	}
	for _, c := range cases {
		typ, err := Parse(c.typ)
		if err != nil {
			t.Fatal(err)
		}
		if got := MatchJSON(c.val, typ); got != c.want {
			t.Errorf("MatchJSON(%s, %v) = %v, want %v", c.typ, c.val, got, c.want)
		}
	}
}
