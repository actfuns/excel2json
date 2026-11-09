package jsonpath

import (
	"reflect"
	"testing"
)

func evalPath(t *testing.T, expr string, root any) []any {
	t.Helper()
	p, err := Parse(expr)
	if err != nil {
		t.Fatalf("Parse(%q): %v", expr, err)
	}
	nodes := p.Eval(root)
	out := make([]any, len(nodes))
	for i, n := range nodes {
		out[i] = n.Value
	}
	return out
}

func TestJSONPathSelectors(t *testing.T) {
	root := map[string]any{
		"store": map[string]any{
			"book": []any{
				map[string]any{"title": "A", "price": 10},
				map[string]any{"title": "B", "price": 20},
			},
		},
		"list": []any{0, 1, 2, 3, 4},
	}

	cases := []struct {
		expr string
		want []any
	}{
		{"$", []any{root}},
		{"$.store.book[0].title", []any{"A"}},
		{"$['store']['book'][1]['title']", []any{"B"}},
		{`$["store"]["book"][-1]["price"]`, []any{20}},
		{"$.store.book[*].title", []any{"A", "B"}},
		{"$.list[0:2]", []any{0, 1}},
		{"$.list[::-1]", []any{4, 3, 2, 1, 0}},
		{"$.list[0::2]", []any{0, 2, 4}},
		{"$.list[1::2]", []any{1, 3}},
		{"$.list[0:-1:2]", []any{0, 2}},
		{"$.list[-100::-1]", []any{}},
		{"$.list[100::-1]", []any{4, 3, 2, 1, 0}},
		{"$.list[4:0:-2]", []any{4, 2}},
		{"$.list[-1:-4:-1]", []any{4, 3, 2}},
		{"$.list[0,2,4]", []any{0, 2, 4}},
		{"$['store']['book'][0]['title']", []any{"A"}},
		{"$['store','list']", []any{map[string]any{"book": []any{
			map[string]any{"title": "A", "price": 10},
			map[string]any{"title": "B", "price": 20},
		}}, []any{0, 1, 2, 3, 4}}},
		{"$..title", []any{"A", "B"}},
		{"$.store..price", []any{10, 20}},
	}
	for _, c := range cases {
		got := evalPath(t, c.expr, root)
		if !reflect.DeepEqual(got, c.want) {
			t.Errorf("%s = %#v, want %#v", c.expr, got, c.want)
		}
	}
}

func TestJSONPathFlatWildcardSlices(t *testing.T) {
	root := []any{
		[]any{"w1", "i1", 10, "w3", "i1", 20},
		[]any{"w2", "i2", 20},
	}
	got := evalPath(t, "$[*][1::3]", root)
	want := []any{"i1", "i1", "i2"}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("$[*][1::3] = %#v, want %#v", got, want)
	}

	got = evalPath(t, "$[*][0]", root)
	want = []any{"w1", "w2"}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("$[*][0] = %#v, want %#v", got, want)
	}
}

func TestJSONPathQuotedNamesWithSeparators(t *testing.T) {
	root := map[string]any{"a,b": 1, "a:b": 2}
	for expr, want := range map[string]any{
		"$['a,b']": 1,
		"$['a:b']": 2,
	} {
		got := evalPath(t, expr, root)
		if len(got) != 1 || got[0] != want {
			t.Errorf("%s = %#v, want %v", expr, got, want)
		}
	}
}

func TestJSONPathInvalid(t *testing.T) {
	bad := []string{
		"", "store", "$.", "$[", "$[]", "$[a]", "$[1:2:0]", "$['x", "$..", "$[1]x",
	}
	for _, expr := range bad {
		if _, err := Parse(expr); err == nil {
			t.Errorf("Parse(%q): expected error", expr)
		}
	}
}

func TestJSONPathSetAtPath(t *testing.T) {
	root := map[string]any{
		"a": []any{1, 2, 3},
	}
	next, err := SetAtPath(root, []Step{
		{Name: "a"},
		{Index: 1, IsIndex: true},
	}, 99)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(next, map[string]any{"a": []any{1, 99, 3}}) {
		t.Errorf("SetAtPath = %#v", next)
	}

	replaced, err := SetAtPath(root, nil, "new")
	if err != nil {
		t.Fatal(err)
	}
	if replaced != "new" {
		t.Errorf("root replacement = %#v", replaced)
	}
}
