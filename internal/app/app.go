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
	root.Flags().StringVar(&opts.Tag, "tag", "", "only export sheets/columns with this tag; untagged = common")
	root.Flags().BoolVar(&opts.StrictTags, "strict-tags", false, "with --tag, error on untagged sheets/columns")
	root.Flags().BoolVar(&opts.Verbose, "verbose", false, "log skipped sheets/columns and files")
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
	root.Flags().BoolVarP(&opts.ListOutput, "array", "a", false, "deprecated: use --list")
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
	opts.Tag = strings.ToLower(strings.TrimSpace(opts.Tag))
	if opts.StrictTags && opts.Tag == "" {
		return fmt.Errorf("--strict-tags requires --tag")
	}

	// Pass 1: load every workbook so ref(...) can resolve across files.
	inputs := make([]loadedInput, len(files))
	loadGroup, loadCtx := errgroup.WithContext(context.Background())
	loadGroup.SetLimit(2 * runtime.NumCPU())
	for i, file := range files {
		loadGroup.Go(func() error {
			select {
			case <-loadCtx.Done():
				return loadCtx.Err()
			default:
			}
			in, err := loadInput(file, opts)
			if err != nil {
				return err
			}
			inputs[i] = in
			return nil
		})
	}
	if err := loadGroup.Wait(); err != nil {
		return err
	}

	contextSheets := make([]convert.SheetContext, 0)
	for _, in := range inputs {
		var engine *convert.Engine
		if in.opts.Verbose {
			engine = convert.New(in.opts)
		}
		for _, sheet := range in.sheets {
			if engine != nil {
				switch {
				case in.opts.Tag != "" && !core.TagMatches(sheet.Tags, in.opts.Tag):
					verbosef("skip sheet %q in %s for output: tag %q not matched\n", sheet.Name, in.path, in.opts.Tag)
				case !engine.HasSelectedColumns(sheet):
					if in.opts.Tag != "" {
						verbosef("skip sheet %q in %s for output: no columns for tag %q\n", sheet.Name, in.path, in.opts.Tag)
					} else {
						verbosef("skip sheet %q in %s for output: no selected columns\n", sheet.Name, in.path)
					}
				}
			}
			name := core.SanitizeName(core.ExpandNameTemplate(in.opts.NameTmpl, in.opts.SourceStem, sheet.Name))
			contextSheets = append(contextSheets, convert.SheetContext{
				Sheet:      sheet,
				OutputName: name,
				TypeDefs:   in.opts.TypeDefs,
			})
		}
	}

	outputNames := make(map[string]string)
	selectedSheets := 0
	for _, in := range inputs {
		selected, err := convert.New(in.opts).FilterSheets(in.sheets)
		if err != nil {
			return fmt.Errorf("%s: %w", in.path, err)
		}
		selectedSheets += len(selected)
		for _, sheet := range in.sheets {
			name := core.SanitizeName(core.ExpandNameTemplate(in.opts.NameTmpl, in.opts.SourceStem, sheet.Name))
			if prev, exists := outputNames[name]; exists {
				return fmt.Errorf("output name collision %q between %s and %s:%s (check --name-tpl)",
					name, prev, in.path, sheet.Name)
			}
			outputNames[name] = in.path + ":" + sheet.Name
		}
	}
	if opts.Tag != "" && selectedSheets == 0 {
		return fmt.Errorf("tag %q matched no sheets or columns", opts.Tag)
	}

	// Pass 2: export each workbook with the shared ref context.
	multi := len(files) > 1
	results := make([]fileResult, len(files))
	exportGroup, exportCtx := errgroup.WithContext(context.Background())
	exportGroup.SetLimit(2 * runtime.NumCPU())
	for i := range inputs {
		exportGroup.Go(func() error {
			select {
			case <-exportCtx.Done():
				return exportCtx.Err()
			default:
			}
			r, err := exportInput(inputs[i], contextSheets, multi)
			if err != nil {
				return err
			}
			results[i] = r
			return nil
		})
	}
	if err := exportGroup.Wait(); err != nil {
		return err
	}

	exported := false
	for _, r := range results {
		if r.exported {
			exported = true
		}
		for _, line := range r.lines {
			fmt.Printf("→ %s\n", line)
		}
	}
	if opts.Tag != "" && !exported {
		return fmt.Errorf("tag %q matched no exportable sheets", opts.Tag)
	}
	return nil
}

func verbosef(format string, args ...any) {
	fmt.Fprintf(os.Stderr, "verbose: "+format, args...)
}

type loadedInput struct {
	path   string
	opts   core.Options
	sheets []core.Sheet
}

func loadInput(file string, opts core.Options) (loadedInput, error) {
	defs, err := schema.LoadTypeDefs(file, opts.TypesPath)
	if err != nil {
		return loadedInput{}, fmt.Errorf("%s: %w", file, err)
	}
	opts.TypeDefs = defs
	opts.SourceStem = core.FileStem(file)
	sheets, err := excel.Load(file, opts)
	if err != nil {
		return loadedInput{}, fmt.Errorf("%s: %w", file, err)
	}
	return loadedInput{path: file, opts: opts, sheets: sheets}, nil
}

type fileResult struct {
	lines    []string
	exported bool
}

func exportInput(in loadedInput, context []convert.SheetContext, multi bool) (fileResult, error) {
	opts := in.opts
	sheets := in.sheets
	if opts.Tag != "" {
		valid, err := convert.NewWithContext(opts, context).FilterSheets(sheets)
		if err != nil {
			return fileResult{}, fmt.Errorf("%s: %w", in.path, err)
		}
		if len(valid) == 0 {
			if opts.Verbose {
				verbosef("skip file %s: no exportable sheets for tag %q\n", in.path, opts.Tag)
			}
			return fileResult{}, nil
		}
		sheets = valid
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
				return fileResult{}, fmt.Errorf("%s: mkdir %s: %w", in.path, target, err)
			}
			path = filepath.Join(target, core.FileStem(in.path)+".json")
		}

		if !opts.MergeSheets && targetIsDir {
			engine := convert.NewWithContext(opts, context)
			valid, err := engine.FilterSheets(sheets)
			if err != nil {
				return fileResult{}, fmt.Errorf("%s: %w", in.path, err)
			}
			if len(valid) == 0 {
				return fileResult{}, fmt.Errorf("%s: no valid sheets to export", in.path)
			}
			if err := os.MkdirAll(target, 0755); err != nil {
				return fileResult{}, fmt.Errorf("%s: mkdir %s: %w", in.path, target, err)
			}
			seen := make(map[string]bool, len(valid))
			for _, sheet := range valid {
				name := core.ExpandNameTemplate(opts.NameTmpl, core.FileStem(in.path), sheet.Name) + ".json"
				spath := filepath.Join(target, core.SanitizeName(name))
				if seen[spath] {
					return fileResult{}, fmt.Errorf("%s: output name collision %q (check --name-tpl)", in.path, spath)
				}
				seen[spath] = true
				if err := writeOneJSONWithEngine(engine, sheet, opts, spath); err != nil {
					return fileResult{}, fmt.Errorf("%s: %w", in.path, err)
				}
				lines = append(lines, spath)
			}
			break
		}

		if err := writeJSONWithContext(sheets, context, opts, path); err != nil {
			return fileResult{}, fmt.Errorf("%s: %w", in.path, err)
		}
		if path != "" {
			lines = append(lines, path)
		}

	case "csv":
		path := target
		if targetIsDir {
			if err := os.MkdirAll(target, 0755); err != nil {
				return fileResult{}, fmt.Errorf("%s: mkdir %s: %w", in.path, target, err)
			}
			if opts.MergeSheets {
				path = filepath.Join(target, core.FileStem(in.path)+".csv")
			}
		}
		csvOpts := opts
		if target != "" && !targetIsDir {
			csvOpts.MergeSheets = true
		}
		paths, err := export.CSVWithContext(sheets, context, csvOpts, path)
		if err != nil {
			return fileResult{}, fmt.Errorf("%s: %w", in.path, err)
		}
		lines = append(lines, paths...)

	default:
		return fileResult{}, fmt.Errorf("unknown format %q (must be json or csv)", opts.Format)
	}

	return fileResult{lines: lines, exported: true}, nil
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
	return writeJSONWithContext(sheets, convert.SheetsContext(opts, sheets), opts, target)
}

func writeJSONWithContext(sheets []core.Sheet, context []convert.SheetContext, opts core.Options, target string) error {
	content, err := export.JSONWithContext(sheets, context, opts)
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
