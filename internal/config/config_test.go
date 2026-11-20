package config

import (
	"os"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/actfuns/excel2json/internal/core"
)

func writeConfig(t *testing.T, dir, body string) string {
	t.Helper()
	path := filepath.Join(dir, "config.yaml")
	if err := os.WriteFile(path, []byte(body), 0644); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestLoadAndResolve(t *testing.T) {
	dir := t.TempDir()
	if err := os.WriteFile(filepath.Join(dir, "a.xlsx"), []byte("x"), 0644); err != nil {
		t.Fatal(err)
	}
	path := writeConfig(t, dir, `
version: 1
inputs:
  - "*.xlsx"
out: dist/
format: json
encoding: utf8-bom
header_row: 0
type_row: 2
data_row: 3
date_format: "yyyy-MM-dd"
parse_cell_json: true
stringify_all: false
lowercase: true
exclude: ["#"]
tag: c
strict_tags: true
verbose: true
required_by_default: false
list: true
merge_sheets: false
name_tpl: "{sheet}"
pretty: true
key: id
unique_keys: true
types: types.yaml
`)
	f, err := Load(path)
	if err != nil {
		t.Fatal(err)
	}
	if f.Version != 1 || f.OutputPath == nil || *f.OutputPath != "dist/" {
		t.Fatalf("unexpected config: %+v", f)
	}
	inputs, err := f.ResolveInputs()
	if err != nil {
		t.Fatal(err)
	}
	if len(inputs) != 1 || !filepath.IsAbs(inputs[0]) {
		t.Errorf("ResolveInputs = %v, want one absolute path", inputs)
	}
	var opts core.Options
	if err := f.ApplyTypes(&opts); err != nil {
		t.Fatal(err)
	}
	if want := filepath.Join(dir, "types.yaml"); opts.TypesPath != want {
		t.Errorf("TypesPath = %q, want %q", opts.TypesPath, want)
	}
}

func TestInlineTypes(t *testing.T) {
	dir := t.TempDir()
	path := writeConfig(t, dir, `
types:
  itemId: "int:gt(0)"
`)
	f, err := Load(path)
	if err != nil {
		t.Fatal(err)
	}
	var opts core.Options
	if err := f.ApplyTypes(&opts); err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(opts.TypeDefs, map[string]string{"itemId": "int:gt(0)"}) {
		t.Errorf("TypeDefs = %v", opts.TypeDefs)
	}
}

func TestUnknownField(t *testing.T) {
	dir := t.TempDir()
	path := writeConfig(t, dir, "unknown_option: true\n")
	if _, err := Load(path); err == nil {
		t.Error("unknown config field: expected error")
	}
}

func TestExampleConfig(t *testing.T) {
	f, err := Load("../../config.example.yaml")
	if err != nil {
		t.Fatal(err)
	}
	if f.Version != 1 {
		t.Errorf("example version = %d, want 1", f.Version)
	}
}

func TestInlineTypesList(t *testing.T) {
	dir := t.TempDir()
	path := writeConfig(t, dir, `
types:
  - name: itemId
    format: "int:gt(0)"
`)
	f, err := Load(path)
	if err != nil {
		t.Fatal(err)
	}
	var opts core.Options
	if err := f.ApplyTypes(&opts); err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(opts.TypeDefs, map[string]string{"itemId": "int:gt(0)"}) {
		t.Errorf("TypeDefs = %v", opts.TypeDefs)
	}
}
