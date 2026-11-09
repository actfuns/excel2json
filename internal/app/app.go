// Package app wires the CLI and the conversion pipeline together.
package app

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"

	"github.com/actfuns/excel2json/internal/convert"
	"github.com/actfuns/excel2json/internal/core"
	"github.com/actfuns/excel2json/internal/excel"
	"github.com/actfuns/excel2json/internal/export"
	"github.com/actfuns/excel2json/internal/schema"
	"github.com/spf13/cobra"
	"golang.org/x/sync/errgroup"
	"golang.org/x/text/encoding/simplifiedchinese"
)

var stdoutMu sync.Mutex

// Main executes the command tree and exits non-zero on error.
func Main(version string) {
	if err := NewCommand(version).Execute(); err != nil {
		fmt.Fprintln(os.Stderr, "excel2json:", err)
		os.Exit(1)
	}
}

// NewCommand builds the root command.
func NewCommand(version string) *cobra.Command {
	var opts core.Options
	var legacySkipRows int
	var legacyName string

	root := &cobra.Command{
		Use:           "excel2json [options] <inputs...>",
		Short:         "Convert Excel (.xls/.xlsx) to JSON / CSV",
		Version:       version,
		SilenceUsage:  true,
		SilenceErrors: true,
		RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) == 0 {
				return cmd.Help()
			}
			if cmd.Flags().Changed("skip-rows") && !cmd.Flags().Changed("data-row") {
				opts.DataRow = legacySkipRows
			}
			if cmd.Flags().Changed("name") && !cmd.Flags().Changed("name-tpl") {
				switch legacyName {
				case "file":
					opts.NameTmpl = "{file}"
				case "sheet":
					opts.NameTmpl = "{sheet}"
				case "both":
					opts.NameTmpl = "{file}_{sheet}"
				}
			}
			return Run(expandInputs(args), opts)
		},
	}

	// Input / output.
	root.Flags().StringVarP(&opts.OutputPath, "out", "o", "", "output path (file or directory)")
	root.Flags().StringVarP(&opts.Format, "format", "f", "json", "output format: json or csv")

	// Data parsing.
	root.Flags().IntVarP(&opts.HeaderRow, "header-row", "n", 0, "row index for column names (0-based)")
	root.Flags().IntVarP(&opts.TypeRow, "type-row", "t", -1, "row index for type annotations (-1 to disable)")
	root.Flags().IntVarP(&opts.DataRow, "data-row", "s", -1, "0-based index of the first data row (-1 = after header/type rows)")
	root.Flags().StringVarP(&opts.Encoding, "encoding", "e", "utf8-nobom", "output encoding: utf8, utf8-nobom, utf8-bom, gbk, gb18030")
	root.Flags().StringVarP(&opts.DateFormat, "date-format", "d", "yyyy/MM/dd", "date format string")

	// Field processing.
	root.Flags().BoolVarP(&opts.Lowercase, "lowercase", "l", false, "convert field names to lowercase")
	root.Flags().StringArrayVarP(&opts.ExcludePrefixes, "exclude", "x", nil, "skip sheets & columns with prefix (repeatable)")
	root.Flags().BoolVar(&opts.ParseCellJSON, "parse-cell-json", false, "parse JSON strings inside cells")
	root.Flags().BoolVar(&opts.StringifyAll, "stringify-all", false, "convert all values to strings")
	root.Flags().BoolVar(&opts.RequireByDefault, "required-by-default", false, "treat fields without !/? as required")

	// Output format.
	root.Flags().BoolVar(&opts.ListOutput, "list", false, "export as a list/array instead of dict")
	root.Flags().BoolVarP(&opts.MergeSheets, "merge-sheets", "m", false, "merge all sheets into one file (default: split)")
	root.Flags().StringVar(&opts.NameTmpl, "name-tpl", "{file}_{sheet}", "output filename template with {file} {sheet}")
	root.Flags().BoolVarP(&opts.Pretty, "pretty", "p", false, "pretty-print JSON")
	root.Flags().StringVarP(&opts.KeyColumn, "key", "k", "", "column name for dict keys (default: first column)")
	root.Flags().BoolVar(&opts.UniqueKeys, "unique-keys", false, "error on duplicate or empty dict keys")
	root.Flags().StringVar(&opts.TypesPath, "types", "", "path to types.yaml/types.yml (auto-discovered next to input or in cwd)")

	// Deprecated aliases kept for compatibility. They are hidden from help.
	root.Flags().BoolVar(&opts.ParseCellJSON, "cell-json", false, "deprecated: use --parse-cell-json")
	_ = root.Flags().MarkDeprecated("cell-json", "use --parse-cell-json instead")
	root.Flags().BoolVar(&opts.StringifyAll, "all-string", false, "deprecated: use --stringify-all")
	_ = root.Flags().MarkDeprecated("all-string", "use --stringify-all instead")
	root.Flags().BoolVar(&opts.ListOutput, "array", false, "deprecated: use --list")
	_ = root.Flags().MarkDeprecated("array", "use --list instead")
	root.Flags().BoolVar(&opts.MergeSheets, "merge", false, "deprecated: use --merge-sheets")
	_ = root.Flags().MarkDeprecated("merge", "use --merge-sheets instead")
	root.Flags().BoolVar(&opts.RequireByDefault, "required", false, "deprecated: use --required-by-default")
	_ = root.Flags().MarkDeprecated("required", "use --required-by-default instead")
	root.Flags().BoolVar(&opts.Lowercase, "lowcase", false, "deprecated: use --lowercase")
	_ = root.Flags().MarkDeprecated("lowcase", "use --lowercase instead")
	root.Flags().IntVar(&opts.HeaderRow, "name-row", 0, "deprecated: use --header-row")
	_ = root.Flags().MarkDeprecated("name-row", "use --header-row instead")
	root.Flags().IntVar(&legacySkipRows, "skip-rows", 1, "deprecated: use --data-row")
	_ = root.Flags().MarkDeprecated("skip-rows", "use --data-row instead")
	root.Flags().StringVar(&legacyName, "name", "", "deprecated: use --name-tpl")
	_ = root.Flags().MarkDeprecated("name", "use --name-tpl instead")

	return root
}

// Run processes every input file concurrently.
func Run(files []string, opts core.Options) error {
	if len(files) == 0 {
		return fmt.Errorf("no input files")
	}
	opts.Format = strings.ToLower(strings.TrimSpace(opts.Format))
	if opts.Format == "" {
		opts.Format = "json"
	}
	if opts.NameTmpl == "" {
		opts.NameTmpl = "{file}_{sheet}"
	}
	if opts.ListOutput && opts.KeyColumn != "" {
		return fmt.Errorf("--key cannot be used with --list")
	}

	multi := len(files) > 1
	g, ctx := errgroup.WithContext(context.Background())
	g.SetLimit(2 * runtime.NumCPU())

	results := make([]fileResult, len(files))
	for i, f := range files {
		g.Go(func() error {
			select {
			case <-ctx.Done():
				return ctx.Err()
			default:
			}
			r, err := exportFile(f, opts, multi)
			if err != nil {
				return err
			}
			results[i] = r
			return nil
		})
	}
	if err := g.Wait(); err != nil {
		return err
	}

	for _, r := range results {
		for _, line := range r.lines {
			fmt.Printf("→ %s\n", line)
		}
	}
	return nil
}

type fileResult struct {
	lines []string
}

func exportFile(f string, opts core.Options, multi bool) (fileResult, error) {
	defs, err := schema.LoadTypeDefs(f, opts.TypesPath)
	if err != nil {
		return fileResult{}, fmt.Errorf("%s: %w", f, err)
	}
	opts.TypeDefs = defs
	opts.SourceStem = core.FileStem(f)

	sheets, err := excel.Load(f, opts)
	if err != nil {
		return fileResult{}, fmt.Errorf("%s: %w", f, err)
	}

	target, targetIsDir := resolveTarget(multi, opts)
	var lines []string

	if target == "" {
		stdoutMu.Lock()
		defer stdoutMu.Unlock()
	}

	switch opts.Format {
	case "json":
		path := target
		if targetIsDir {
			if err := os.MkdirAll(target, 0755); err != nil {
				return fileResult{}, fmt.Errorf("%s: mkdir %s: %w", f, target, err)
			}
			path = filepath.Join(target, core.FileStem(f)+".json")
		}

		if !opts.MergeSheets && targetIsDir {
			engine := convert.NewWithSheets(opts, sheets)
			valid := engine.FilterSheets(sheets)
			if len(valid) == 0 {
				return fileResult{}, fmt.Errorf("%s: no valid sheets to export", f)
			}
			dir := target
			if err := os.MkdirAll(dir, 0755); err != nil {
				return fileResult{}, fmt.Errorf("%s: mkdir %s: %w", f, dir, err)
			}
			seen := make(map[string]bool, len(valid))
			for _, s := range valid {
				name := core.ExpandNameTemplate(opts.NameTmpl, core.FileStem(f), s.Name) + ".json"
				spath := filepath.Join(dir, core.SanitizeName(name))
				if seen[spath] {
					return fileResult{}, fmt.Errorf("%s: output name collision %q (check --name-tpl)", f, spath)
				}
				seen[spath] = true
				if err := writeOneJSONWithEngine(engine, s, opts, spath); err != nil {
					return fileResult{}, fmt.Errorf("%s: %w", f, err)
				}
				lines = append(lines, spath)
			}
			break
		}

		if err := writeJSON(sheets, opts, path); err != nil {
			return fileResult{}, fmt.Errorf("%s: %w", f, err)
		}
		if path != "" {
			lines = append(lines, path)
		}

	case "csv":
		path := target
		if targetIsDir {
			if err := os.MkdirAll(target, 0755); err != nil {
				return fileResult{}, fmt.Errorf("%s: mkdir %s: %w", f, target, err)
			}
			path = filepath.Join(target, core.FileStem(f)+".csv")
		}
		csvOpts := opts
		if target != "" && !targetIsDir {
			csvOpts.MergeSheets = true
		}
		paths, err := export.CSV(sheets, csvOpts, path)
		if err != nil {
			return fileResult{}, fmt.Errorf("%s: %w", f, err)
		}
		lines = append(lines, paths...)

	default:
		return fileResult{}, fmt.Errorf("unknown format %q (must be json or csv)", opts.Format)
	}

	return fileResult{lines: lines}, nil
}

func resolveTarget(multi bool, opts core.Options) (string, bool) {
	if opts.OutputPath == "" {
		return "", false
	}
	if core.IsDirPath(opts.OutputPath) || multi || filepath.Ext(opts.OutputPath) == "" {
		return opts.OutputPath, true
	}
	return opts.OutputPath, false
}

func writeJSON(sheets []core.Sheet, opts core.Options, target string) error {
	content, err := export.JSON(sheets, opts)
	if err != nil {
		return err
	}
	if target == "" {
		_, err := fmt.Println(content)
		return err
	}
	data, err := encodeContent(content, opts.Encoding)
	if err != nil {
		return err
	}
	return os.WriteFile(target, data, 0644)
}

func writeOneJSON(sheet core.Sheet, context []core.Sheet, opts core.Options, path string) error {
	return writeOneJSONWithEngine(convert.NewWithSheets(opts, context), sheet, opts, path)
}

func writeOneJSONWithEngine(engine *convert.Engine, sheet core.Sheet, opts core.Options, path string) error {
	value, err := engine.ConvertSheet(sheet)
	if err != nil {
		return err
	}
	content, err := export.Serialize(value, opts.Pretty)
	if err != nil {
		return err
	}
	data, err := encodeContent(content, opts.Encoding)
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0644)
}

func expandInputs(args []string) []string {
	var out []string
	seen := make(map[string]bool)
	for _, a := range args {
		fi, err := os.Stat(a)
		if err != nil || !fi.IsDir() {
			if !seen[a] {
				seen[a] = true
				out = append(out, a)
			}
			continue
		}
		entries, _ := os.ReadDir(a)
		for _, e := range entries {
			if e.IsDir() {
				continue
			}
			name := strings.ToLower(e.Name())
			if !strings.HasSuffix(name, ".xls") && !strings.HasSuffix(name, ".xlsx") {
				continue
			}
			p := filepath.Join(a, e.Name())
			if !seen[p] {
				seen[p] = true
				out = append(out, p)
			}
		}
	}
	return out
}

func encodeContent(s, enc string) ([]byte, error) {
	switch strings.ToLower(enc) {
	case "", "utf8", "utf8-nobom":
		return []byte(s), nil
	case "utf8-bom":
		return append([]byte{0xEF, 0xBB, 0xBF}, []byte(s)...), nil
	case "gbk":
		return simplifiedchinese.GBK.NewEncoder().Bytes([]byte(s))
	case "gb18030":
		return simplifiedchinese.GB18030.NewEncoder().Bytes([]byte(s))
	default:
		return nil, fmt.Errorf("unknown encoding %q (supported: utf8, utf8-bom, utf8-nobom, gbk, gb18030)", enc)
	}
}
