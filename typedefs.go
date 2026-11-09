package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"

	"gopkg.in/yaml.v3"
)

// Type dictionaries are written as YAML.
//
// Map form:
//
//	types:
//	  itemId: "int:gt(0),ref(items.id)"
//	  items:  "$[*][1::3]:string:ref(item.id)"
//
// List form:
//
//	types:
//	  - name: itemId
//	    format: "int:gt(0),ref(items.id)"
//
// A top-level mapping of name -> format is also accepted.
type yamlTypeEntry struct {
	Name   string `yaml:"name"`
	Format string `yaml:"format"`
}

type cachedTypeDefs struct {
	defs map[string]string
	err  error
}

var typeDefsCache sync.Map // absolute path -> cachedTypeDefs

// loadTypeDefs resolves and loads the type dictionary for one input file.
// An explicit --types path wins; otherwise types.yaml/types.yml are looked up
// next to the input file and then in the current working directory.
func loadTypeDefs(inputFile, explicit string) (map[string]string, error) {
	path, err := resolveTypesPath(inputFile, explicit)
	if err != nil || path == "" {
		return nil, err
	}

	abs, err := filepath.Abs(path)
	if err != nil {
		abs = path
	}
	if v, ok := typeDefsCache.Load(abs); ok {
		c := v.(cachedTypeDefs)
		return c.defs, c.err
	}

	defs, err := readTypeDefs(abs)
	typeDefsCache.Store(abs, cachedTypeDefs{defs: defs, err: err})
	return defs, err
}

func resolveTypesPath(inputFile, explicit string) (string, error) {
	if explicit != "" {
		fi, err := os.Stat(explicit)
		if err != nil {
			return "", fmt.Errorf("types file %s: %w", explicit, err)
		}
		if fi.IsDir() {
			return "", fmt.Errorf("types file %s is a directory", explicit)
		}
		return explicit, nil
	}

	names := []string{"types.yaml", "types.yml"}
	var candidates []string
	if inputFile != "" {
		for _, name := range names {
			candidates = append(candidates, filepath.Join(filepath.Dir(inputFile), name))
		}
	}
	candidates = append(candidates, names...)
	for _, path := range candidates {
		if fi, err := os.Stat(path); err == nil && !fi.IsDir() {
			return path, nil
		}
	}
	return "", nil
}

func readTypeDefs(path string) (map[string]string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("read types file %s: %w", path, err)
	}
	return parseYAMLTypeDefs(path, data)
}

func parseYAMLTypeDefs(path string, data []byte) (map[string]string, error) {
	// "types" as a mapping.
	var wrapped struct {
		Types map[string]string `yaml:"types"`
	}
	if err := yaml.Unmarshal(data, &wrapped); err == nil && len(wrapped.Types) > 0 {
		return normalizeTypeDefs(path, wrapped.Types)
	}

	// "types" as a list of {name, format}.
	var list struct {
		Types []yamlTypeEntry `yaml:"types"`
	}
	if err := yaml.Unmarshal(data, &list); err == nil && len(list.Types) > 0 {
		raw := make(map[string]string, len(list.Types))
		for _, t := range list.Types {
			name := strings.TrimSpace(t.Name)
			if _, exists := raw[name]; exists {
				return nil, fmt.Errorf("%s: duplicate type %q", path, name)
			}
			raw[name] = t.Format
		}
		return normalizeTypeDefs(path, raw)
	}

	// A top-level mapping of name -> format.
	var direct map[string]string
	if err := yaml.Unmarshal(data, &direct); err != nil {
		return nil, fmt.Errorf("parse types file %s: %w", path, err)
	}
	return normalizeTypeDefs(path, direct)
}

func normalizeTypeDefs(path string, raw map[string]string) (map[string]string, error) {
	if len(raw) == 0 {
		return nil, fmt.Errorf("%s: no type definitions", path)
	}
	defs := make(map[string]string, len(raw))
	for name, format := range raw {
		name = strings.TrimSpace(name)
		format = strings.TrimSpace(format)
		if name == "" {
			return nil, fmt.Errorf("%s: type definition with empty name", path)
		}
		if format == "" {
			return nil, fmt.Errorf("%s: type %q is missing a format", path, name)
		}
		if _, exists := defs[name]; exists {
			return nil, fmt.Errorf("%s: duplicate type %q", path, name)
		}
		if _, err := parseType(name); err == nil {
			return nil, fmt.Errorf("%s: type name %q conflicts with a built-in type", path, name)
		}
		defs[name] = format
	}
	return defs, nil
}
