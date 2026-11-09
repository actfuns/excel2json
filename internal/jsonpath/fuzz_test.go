package jsonpath

import "testing"

func FuzzParse(f *testing.F) {
	for _, s := range []string{
		"$", "$.a", "$[0]", "$[-1]", "$[*]", "$[0:2]", "$[1::3]",
		"$['a']", `$["a"]`, "$..a", "$[*][1::3]", "$[0,2]",
	} {
		f.Add(s)
	}
	f.Fuzz(func(t *testing.T, src string) {
		p, err := Parse(src)
		if err != nil {
			return
		}
		_ = p.Eval([]any{map[string]any{"a": []any{1, 2, 3}}})
		_ = p.Eval([]any{1, 2, 3})
	})
}
