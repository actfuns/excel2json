package main

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestWriteCSVSeq(t *testing.T) {
	s1 := Sheet{Name: "S", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"1"}}}
	s2 := Sheet{Name: "T", Headers: []string{"id"}, Types: []string{"int"}, Rows: [][]string{{"2"}}}

	var buf bytes.Buffer
	if err := writeCSVSeq(&buf, []Sheet{s1, s2}, NewExporter(Options{})); err != nil {
		t.Fatal(err)
	}
	want := "id\n1\n--- T ---\nid\n2\n"
	if buf.String() != want {
		t.Errorf("writeCSVSeq = %q, want %q", buf.String(), want)
	}
}

func TestExportCSVFiles(t *testing.T) {
	s1 := Sheet{Name: "S", Headers: []string{"id", "v"}, Types: []string{"int", "string"}, Rows: [][]string{{"1", "a"}}}
	s2 := Sheet{Name: "T", Headers: []string{"id", "v"}, Types: []string{"int", "string"}, Rows: [][]string{{"2", "b"}}}
	opts := Options{}

	dir := t.TempDir()
	single := filepath.Join(dir, "single.csv")
	if err := exportCSV([]Sheet{s1}, opts, single); err != nil {
		t.Fatal(err)
	}
	assertFile(t, single, "id,v\n1,a\n")

	stem := filepath.Join(dir, "multi.csv")
	if err := exportCSV([]Sheet{s1, s2}, opts, stem); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(dir, "multi_S.csv"), "id,v\n1,a\n")
	assertFile(t, filepath.Join(dir, "multi_T.csv"), "id,v\n2,b\n")

	outDir := filepath.Join(dir, "out")
	if err := os.MkdirAll(outDir, 0755); err != nil {
		t.Fatal(err)
	}
	if err := exportCSV([]Sheet{s1, s2}, opts, outDir); err != nil {
		t.Fatal(err)
	}
	assertFile(t, filepath.Join(outDir, "S.csv"), "id,v\n1,a\n")
	assertFile(t, filepath.Join(outDir, "T.csv"), "id,v\n2,b\n")

	if err := exportCSV(nil, opts, ""); err == nil {
		t.Error("exportCSV with no valid sheets: expected error")
	}
}

func assertFile(t *testing.T, path, want string) {
	t.Helper()
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if strings.ReplaceAll(string(b), "\r\n", "\n") != want {
		t.Errorf("%s = %q, want %q", path, b, want)
	}
}
