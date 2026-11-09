package core

import (
	"os"
	"path/filepath"
	"strings"
)

// FileStem returns the filename without its extension.
func FileStem(path string) string {
	name := filepath.Base(path)
	if i := strings.LastIndex(name, "."); i > 0 {
		return name[:i]
	}
	return name
}

// ExpandNameTemplate expands the {file} and {sheet} placeholders.
func ExpandNameTemplate(tmpl, file, sheet string) string {
	s := strings.ReplaceAll(tmpl, "{file}", file)
	return strings.ReplaceAll(s, "{sheet}", sheet)
}

// SanitizeName replaces characters that are unsafe in file names.
func SanitizeName(name string) string {
	r := strings.NewReplacer(
		"/", "_", "\\", "_", ":", "_", "*", "_",
		"?", "_", "\"", "_", "<", "_", ">", "_", "|", "_", " ", "_",
	)
	name = strings.TrimRight(r.Replace(name), ". ")
	if name == "" {
		return "_"
	}
	base := name
	if i := strings.IndexByte(base, '.'); i >= 0 {
		base = base[:i]
	}
	switch strings.ToUpper(base) {
	case "CON", "PRN", "AUX", "NUL",
		"COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
		"LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9":
		name = "_" + name
	}
	return name
}

// IsDirPath reports whether path looks like a directory.
func IsDirPath(path string) bool {
	if strings.HasSuffix(path, "/") || strings.HasSuffix(path, "\\") {
		return true
	}
	fi, err := os.Stat(path)
	return err == nil && fi.IsDir()
}
