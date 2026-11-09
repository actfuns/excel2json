package core

import (
	"os"
	"path/filepath"
	"testing"
)

func TestFileStem(t *testing.T) {
	cases := map[string]string{
		"a.xlsx":   "a",
		"a.b.c":    "a.b",
		"noext":    "noext",
		".hidden":  ".hidden",
		"dir/x.go": "x",
	}
	for in, want := range cases {
		if got := FileStem(in); got != want {
			t.Errorf("FileStem(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestExpandNameTemplate(t *testing.T) {
	if got := ExpandNameTemplate("my_{file}_{sheet}", "data", "Sheet1"); got != "my_data_Sheet1" {
		t.Errorf("ExpandNameTemplate = %q", got)
	}
}

func TestSanitizeNameAndDirPath(t *testing.T) {
	if got, want := SanitizeName(`a/b\c:d*e?f"g<h>i|j k`), "a_b_c_d_e_f_g_h_i_j_k"; got != want {
		t.Errorf("SanitizeName = %q, want %q", got, want)
	}

	dir := t.TempDir()
	if !IsDirPath(dir) || !IsDirPath(dir+"/") {
		t.Error("IsDirPath should be true for a directory and a trailing slash")
	}
	file := filepath.Join(dir, "x.json")
	if err := os.WriteFile(file, []byte("{}"), 0644); err != nil {
		t.Fatal(err)
	}
	if IsDirPath(file) {
		t.Error("IsDirPath should be false for a file")
	}
	if IsDirPath(filepath.Join(dir, "missing")) {
		t.Error("IsDirPath should be false for a missing path")
	}
}

func TestSplitFieldMarker(t *testing.T) {
	cases := []struct {
		in       string
		want     string
		nullable *bool
	}{
		{"id", "id", nil},
		{"id!", "id", boolPtr(true)},
		{"id?", "id", boolPtr(false)},
	}
	for _, c := range cases {
		got, req := SplitFieldMarker(c.in)
		if got != c.want {
			t.Errorf("SplitFieldMarker(%q) = %q, want %q", c.in, got, c.want)
		}
		if (c.nullable == nil) != (req == nil) || (c.nullable != nil && *c.nullable != *req) {
			t.Errorf("SplitFieldMarker(%q) marker = %v, want %v", c.in, req, c.nullable)
		}
	}
}

func boolPtr(v bool) *bool { return &v }

func TestSanitizeNameReserved(t *testing.T) {
	if got := SanitizeName("CON"); got != "_CON" {
		t.Errorf("SanitizeName(CON) = %q", got)
	}
	if got := SanitizeName("name."); got != "name" {
		t.Errorf("SanitizeName(name.) = %q", got)
	}
}
