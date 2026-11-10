package convert

import (
	"testing"

	"github.com/actfuns/excel2json/internal/core"
)

func FuzzConvertCell(f *testing.F) {
	for _, c := range [][2]string{
		{"int", "42"},
		{"int:gt(0)", "1"},
		{"$[0]:int", "[1]"},
		{"[4]int32:pad,gte(0)", "[1,2]"},
		{"when(typeof($) = string) string:len(1)", "abc"},
		{"string:or(and(email,len(6,64)),phone)", "a@b.com"},
		{"$.a:int", `{"a":5}`},
		{"when(type=1) int:gt(0)", "5"},
	} {
		f.Add(c[0], c[1])
	}
	e := New(core.Options{})
	f.Fuzz(func(t *testing.T, annotation, cell string) {
		_, _ = e.convertCell(annotation, cell)
	})
}
