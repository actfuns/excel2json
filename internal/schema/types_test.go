package schema

import (
	"os"
	"path/filepath"
	"testing"
)

func writeFile(t *testing.T, dir, name, content string) string {
	t.Helper()
	path := filepath.Join(dir, name)
	if err := os.WriteFile(path, []byte(content), 0644); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestLoadTypeDefsAutoDiscover(t *testing.T) {
	dir := t.TempDir()
	writeFile(t, dir, "types.yaml", `
types:
  itemId: "int:gt(0)"
  items: "$[*][1::3]:string:ref(item.id)"
`)
	input := filepath.Join(dir, "data.xlsx")

	defs, err := LoadTypeDefs(input, "")
	if err != nil {
		t.Fatal(err)
	}
	if got := defs["itemId"]; got != "int:gt(0)" {
		t.Errorf("itemId = %q", got)
	}
	if got := defs["items"]; got != "$[*][1::3]:string:ref(item.id)" {
		t.Errorf("items = %q", got)
	}
}

func TestLoadTypeDefsYAMLForms(t *testing.T) {
	dir := t.TempDir()

	listPath := writeFile(t, dir, "list.yml", `
types:
  - name: score
    format: "int:between(0,100)"
  - name: name
    format: "string:len(1,20)"
`)
	defs, err := LoadTypeDefs("", listPath)
	if err != nil {
		t.Fatal(err)
	}
	if defs["score"] != "int:between(0,100)" || defs["name"] != "string:len(1,20)" {
		t.Errorf("list form = %v", defs)
	}

	directPath := writeFile(t, dir, "direct.yaml", `
score: "int:between(0,100)"
`)
	defs, err = LoadTypeDefs("", directPath)
	if err != nil {
		t.Fatal(err)
	}
	if defs["score"] != "int:between(0,100)" {
		t.Errorf("direct form = %v", defs)
	}
}

func TestLoadTypeDefsYAMLPrecedence(t *testing.T) {
	dir := t.TempDir()
	writeFile(t, dir, "types.yml", `
types:
  items: "int"
`)
	writeFile(t, dir, "types.yaml", `
types:
  items: "string"
`)
	input := filepath.Join(dir, "data.xlsx")

	defs, err := LoadTypeDefs(input, "")
	if err != nil {
		t.Fatal(err)
	}
	if defs["items"] != "string" {
		t.Errorf("types.yaml should win, got %q", defs["items"])
	}
}

func TestLoadTypeDefsExplicitAndParse(t *testing.T) {
	dir := t.TempDir()
	path := writeFile(t, dir, "custom.yaml", `
types:
  score: "int:between(0,100)"
`)

	defs, err := LoadTypeDefs("", path)
	if err != nil {
		t.Fatal(err)
	}
	spec, err := Parse("#score", defs)
	if err != nil {
		t.Fatalf("parse #score: %v", err)
	}
	if spec.Rules[0].Type.String() != "int" {
		t.Errorf("score type = %s", spec.Rules[0].Type)
	}
}

func TestNamedTypeWithoutHash(t *testing.T) {
	defs := map[string]string{
		"itemId": "int:gte(0)",
		"items":  `$[*][1::3]:string:ref(item.id)`,
	}
	for _, annotation := range []string{"itemId", "#itemId"} {
		spec, err := Parse(annotation, defs)
		if err != nil {
			t.Errorf("%s: %v", annotation, err)
			continue
		}
		if spec.Rules[0].Type.String() != "int" {
			t.Errorf("%s type = %s", annotation, spec.Rules[0].Type)
		}
	}
}

func TestLoadTypeDefsErrors(t *testing.T) {
	dir := t.TempDir()

	if _, err := LoadTypeDefs("", filepath.Join(dir, "missing.yaml")); err == nil {
		t.Error("missing explicit file: expected error")
	}

	dup := writeFile(t, dir, "dup.yaml", `
types:
  - name: a
    format: "int"
  - name: a
    format: "string"
`)
	if _, err := LoadTypeDefs("", dup); err == nil {
		t.Error("duplicate name: expected error")
	}

	empty := writeFile(t, dir, "empty.yaml", `
types:
  - name: a
`)
	if _, err := LoadTypeDefs("", empty); err == nil {
		t.Error("missing format: expected error")
	}

	bad := writeFile(t, dir, "bad.yaml", `types: [`)
	if _, err := LoadTypeDefs("", bad); err == nil {
		t.Error("malformed yaml: expected error")
	}

	builtin := writeFile(t, dir, "builtin.yaml", `
types:
  int: "string"
`)
	if _, err := LoadTypeDefs("", builtin); err == nil {
		t.Error("built-in type name collision: expected error")
	}

	badName := writeFile(t, dir, "badname.yaml", `
types:
  "bad name": "int"
`)
	if _, err := LoadTypeDefs("", badName); err == nil {
		t.Error("invalid type name: expected error")
	}
}
