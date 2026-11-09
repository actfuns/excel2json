package typing

import "testing"

func FuzzParse(f *testing.F) {
	for _, s := range []string{
		"int", "int8", "[]int", "[3]int", "map[string]any", "date", "list", "object", "int!", "int?",
	} {
		f.Add(s)
	}
	f.Fuzz(func(t *testing.T, src string) {
		typ, err := Parse(src)
		if err != nil {
			return
		}
		c := &Converter{DateFormat: "yyyy/MM/dd"}
		_, _ = c.ConvertString("1", typ)
		v, err := DecodeJSONCell(`[1,"a",true]`)
		if err == nil {
			_, _ = c.ConvertJSON(v, typ)
		}
	})
}
