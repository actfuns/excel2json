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

func TestSplitTagSuffix(t *testing.T) {
	cases := []struct {
		in   string
		want string
		tags []string
		ok   bool
	}{
		{"reward", "reward", nil, false},
		{"reward@c", "reward", []string{"c"}, true},
		{"reward@s,c", "reward", []string{"s", "c"}, true},
		{"reward@C", "reward", []string{"c"}, true},
		{"reward@", "reward@", nil, false},
		{"reward@c,d.e", "reward@c,d.e", nil, false},
		{"reward~c", "reward~c", nil, false},
		{"@c", "@c", nil, false},
	}
	for _, c := range cases {
		got, tags, ok := SplitTagSuffix(c.in)
		if got != c.want || ok != c.ok {
			t.Errorf("SplitTagSuffix(%q) = %q,%v,%v, want %q,%v,%v", c.in, got, tags, ok, c.want, c.tags, c.ok)
			continue
		}
		if len(tags) != len(c.tags) {
			t.Errorf("SplitTagSuffix(%q) tags = %v, want %v", c.in, tags, c.tags)
			continue
		}
		for i := range tags {
			if tags[i] != c.tags[i] {
				t.Errorf("SplitTagSuffix(%q) tags = %v, want %v", c.in, tags, c.tags)
				break
			}
		}
	}
}

func TestSplitHeader(t *testing.T) {
	cases := []struct {
		in       string
		markers  bool
		name     string
		required *bool
		tags     []string
	}{
		{"reward@c", true, "reward", nil, []string{"c"}},
		{"reward@c!", true, "reward", boolPtr(true), []string{"c"}},
		{"reward!@c", true, "reward", boolPtr(true), []string{"c"}},
		{"reward@c?", true, "reward", boolPtr(false), []string{"c"}},
		{"reward!@c", false, "reward!", nil, []string{"c"}},
		{"plain", true, "plain", nil, nil},
		{"reward @c", true, "reward", nil, []string{"c"}},
	}
	for _, c := range cases {
		name, req, tags := SplitHeader(c.in, c.markers)
		if name != c.name {
			t.Errorf("SplitHeader(%q) name = %q, want %q", c.in, name, c.name)
		}
		if (req == nil) != (c.required == nil) || (req != nil && *req != *c.required) {
			t.Errorf("SplitHeader(%q) required = %v, want %v", c.in, req, c.required)
		}
		if len(tags) != len(c.tags) {
			t.Errorf("SplitHeader(%q) tags = %v, want %v", c.in, tags, c.tags)
		}
	}
}

func TestTagMatches(t *testing.T) {
	if !TagMatches(nil, "c") || !TagMatches([]string{"c"}, "") {
		t.Error("empty tags or target should match")
	}
	if !TagMatches([]string{"C", "s"}, "c") {
		t.Error("tag matching should be case-insensitive")
	}
	if TagMatches([]string{"s"}, "c") {
		t.Error("s should not match c")
	}
}
