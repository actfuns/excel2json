package main

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

	defs, err := loadTypeDefs(input, "")
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
	defs, err := loadTypeDefs("", listPath)
	if err != nil {
		t.Fatal(err)
	}
	if defs["score"] != "int:between(0,100)" || defs["name"] != "string:len(1,20)" {
		t.Errorf("list form = %v", defs)
	}

	directPath := writeFile(t, dir, "direct.yaml", `
score: "int:between(0,100)"
`)
	defs, err = loadTypeDefs("", directPath)
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

	defs, err := loadTypeDefs(input, "")
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

	defs, err := loadTypeDefs("", path)
	if err != nil {
		t.Fatal(err)
	}
	e := NewExporter(Options{TypeDefs: defs})
	if _, err := e.convertCell("#score", "50"); err != nil {
		t.Fatalf("score 50: %v", err)
	}
	if _, err := e.convertCell("#score", "101"); err == nil {
		t.Error("score 101: expected error")
	}
}

func TestNamedTypeWithoutHash(t *testing.T) {
	defs := map[string]string{
		"itemId": "int:gte(0)",
		"items":  `$[*][1::3]:string:ref(item.id)`,
	}
	e := NewExporter(Options{TypeDefs: defs})

	for _, annotation := range []string{"itemId", "#itemId"} {
		v, err := e.convertCell(annotation, "5")
		if err != nil {
			t.Errorf("%s: %v", annotation, err)
			continue
		}
		if got := marshalValue(t, v); got != "5" {
			t.Errorf("%s = %s, want 5", annotation, got)
		}
	}
}

func TestLoadTypeDefsErrors(t *testing.T) {
	dir := t.TempDir()

	if _, err := loadTypeDefs("", filepath.Join(dir, "missing.yaml")); err == nil {
		t.Error("missing explicit file: expected error")
	}

	dup := writeFile(t, dir, "dup.yaml", `
types:
  - name: a
    format: "int"
  - name: a
    format: "string"
`)
	if _, err := loadTypeDefs("", dup); err == nil {
		t.Error("duplicate name: expected error")
	}

	empty := writeFile(t, dir, "empty.yaml", `
types:
  - name: a
`)
	if _, err := loadTypeDefs("", empty); err == nil {
		t.Error("missing format: expected error")
	}

	bad := writeFile(t, dir, "bad.yaml", `types: [`)
	if _, err := loadTypeDefs("", bad); err == nil {
		t.Error("malformed yaml: expected error")
	}

	builtin := writeFile(t, dir, "builtin.yaml", `
types:
  int: "string"
`)
	if _, err := loadTypeDefs("", builtin); err == nil {
		t.Error("built-in type name collision: expected error")
	}
}
