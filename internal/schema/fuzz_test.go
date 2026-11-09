package schema

import "testing"

func FuzzParse(f *testing.F) {
	for _, s := range []string{
		"int", "int:gt(0),lt(10)", "int?:lt(1);string?:minlen(1)",
		"$[*][1::3]:string:ref(item.id)", "[4]int32:pad,gte(0)",
		"string:or(and(email,len(6,64)),phone)", "int:not(gt(0),lt(10))",
	} {
		f.Add(s)
	}
	f.Fuzz(func(t *testing.T, src string) {
		_, _ = Parse(src, map[string]string{"itemId": "int:gt(0)"})
	})
}
