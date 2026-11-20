// Package config loads YAML configuration files for the CLI.
package config

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/actfuns/excel2json/internal/core"
	"github.com/actfuns/excel2json/internal/schema"
	"gopkg.in/yaml.v3"
)

// File is a YAML configuration file. Pointer fields distinguish "not set"
// from explicit zero values so CLI flags can override only what the user set.
type File struct {
	Version int `yaml:"version"`

	Inputs []string `yaml:"inputs"`

	OutputPath *string `yaml:"out"`
	Format     *string `yaml:"format"`
	Encoding   *string `yaml:"encoding"`

	HeaderRow  *int    `yaml:"header_row"`
	TypeRow    *int    `yaml:"type_row"`
	DataRow    *int    `yaml:"data_row"`
	DateFormat *string `yaml:"date_format"`

	ParseCellJSON    *bool    `yaml:"parse_cell_json"`
	StringifyAll     *bool    `yaml:"stringify_all"`
	LowerCase        *bool    `yaml:"lowercase"`
	ExcludePrefixes  []string `yaml:"exclude"`
	Tag              *string  `yaml:"tag"`
	StrictTags       *bool    `yaml:"strict_tags"`
	Verbose          *bool    `yaml:"verbose"`
	RequireByDefault *bool    `yaml:"required_by_default"`

	ListOutput  *bool   `yaml:"list"`
	MergeSheets *bool   `yaml:"merge_sheets"`
	NameTmpl    *string `yaml:"name_tpl"`
	Pretty      *bool   `yaml:"pretty"`
	KeyColumn   *string `yaml:"key"`
	UniqueKeys  *bool   `yaml:"unique_keys"`

	Types yaml.Node `yaml:"types"`

	Path string `yaml:"-"`
	Dir  string `yaml:"-"`
}

// Load reads a config file and records its directory for relative paths.
func Load(path string) (*File, error) {
	abs, err := filepath.Abs(path)
	if err != nil {
		return nil, fmt.Errorf("config %s: %w", path, err)
	}
	data, err := os.ReadFile(abs)
	if err != nil {
		return nil, fmt.Errorf("config %s: %w", path, err)
	}
	dec := yaml.NewDecoder(bytes.NewReader(data))
	dec.KnownFields(true)
	var f File
	if err := dec.Decode(&f); err != nil {
		return nil, fmt.Errorf("config %s: %w", path, err)
	}
	if f.Version != 0 && f.Version != 1 {
		return nil, fmt.Errorf("config %s: unsupported version %d", path, f.Version)
	}
	f.Path = abs
	f.Dir = filepath.Dir(abs)
	return &f, nil
}

// ResolvePath resolves a path relative to the config file.
func (f *File) ResolvePath(path string) string {
	if path == "" || filepath.IsAbs(path) {
		return path
	}
	return filepath.Join(f.Dir, path)
}

// ResolveInputs expands config input paths/globs relative to the config file.
func (f *File) ResolveInputs() ([]string, error) {
	out := make([]string, 0, len(f.Inputs))
	for _, pattern := range f.Inputs {
		pattern = f.ResolvePath(pattern)
		if !strings.ContainsAny(pattern, "*?[") {
			out = append(out, pattern)
			continue
		}
		matches, err := filepath.Glob(pattern)
		if err != nil {
			return nil, fmt.Errorf("config input %q: %w", pattern, err)
		}
		if len(matches) == 0 {
			return nil, fmt.Errorf("config input %q matched no files", pattern)
		}
		out = append(out, matches...)
	}
	return out, nil
}

// ApplyTypes applies inline or referenced type definitions to opts.
func (f *File) ApplyTypes(opts *core.Options) error {
	if f.Types.Kind == 0 || (f.Types.Kind == yaml.ScalarNode && f.Types.Tag == "!!null") {
		return nil
	}
	if f.Types.Kind == yaml.ScalarNode {
		if f.Types.Tag != "!!str" || strings.TrimSpace(f.Types.Value) == "" {
			return fmt.Errorf("config types must be a path or mapping")
		}
		opts.TypesPath = f.ResolvePath(f.Types.Value)
		return nil
	}
	if f.Types.Kind != yaml.MappingNode && f.Types.Kind != yaml.SequenceNode {
		return fmt.Errorf("config types must be a path, mapping or list")
	}
	root := yaml.Node{
		Kind: yaml.MappingNode,
		Content: []*yaml.Node{
			{Kind: yaml.ScalarNode, Tag: "!!str", Value: "types"},
			&f.Types,
		},
	}
	data, err := yaml.Marshal(&root)
	if err != nil {
		return fmt.Errorf("config types: %w", err)
	}
	defs, err := schema.ParseTypeDefs(f.Path, data)
	if err != nil {
		return err
	}
	opts.TypeDefs = defs
	return nil
}
