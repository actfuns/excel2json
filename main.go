package main

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"time"

	"github.com/spf13/cobra"
	"golang.org/x/sync/errgroup"
)

var (
	opts     Options
	stdoutMu sync.Mutex // guards stdout output across concurrent goroutines
)

var rootCmd = &cobra.Command{
	Use:   "excel2json [options] <inputs...>",
	Short: "Convert Excel (.xls/.xlsx) to JSON / CSV",
	Long: `excel2json — Excel to JSON/CSV converter.

Reads one or more Excel files (or directories of .xls/.xlsx files)
and exports each to JSON or CSV.

Output routing (determined by -o and input count):

  excel2json data.xlsx                  → JSON to stdout
  excel2json data.xlsx --format csv     → CSV  to stdout
  excel2json data.xlsx -o out.json      → JSON to out.json
  excel2json data.xlsx -o ./dir/        → ./dir/data.json
  excel2json a.xlsx b.xlsx -o ./dir/    → ./dir/a.json ./dir/b.json`,
	Args: cobra.MinimumNArgs(1),
	RunE: func(cmd *cobra.Command, args []string) error {
		files := expandInputs(args)

		start := time.Now()

		if err := run(files, opts); err != nil {
			return err
		}

		fmt.Printf("Conversion complete in [%dms].\n", time.Since(start).Milliseconds())
		return nil
	},
}

func init() {
	// Output
	rootCmd.Flags().StringVarP(&opts.OutputPath, "out", "o", "", "output path (file for single input, dir for multiple inputs)")
	rootCmd.Flags().StringVar(&opts.Format, "format", "json", "output format: json or csv")

	// Data parsing
	rootCmd.Flags().IntVar(&opts.HeaderRows, "header", 1, "number of header rows")
	rootCmd.Flags().StringVarP(&opts.Encoding, "encoding", "c", "utf8-nobom", "output file encoding")
	rootCmd.Flags().StringVarP(&opts.DateFormat, "date", "d", "yyyy/MM/dd", "date format string")

	// Field processing
	rootCmd.Flags().BoolVarP(&opts.Lowcase, "lowcase", "l", false, "convert field names to lowercase")
	rootCmd.Flags().StringVarP(&opts.ExcludePrefix, "exclude_prefix", "x", "", "skip sheets/columns with this prefix")
	rootCmd.Flags().BoolVar(&opts.CellJSON, "cell_json", false, "parse JSON strings inside cells")
	rootCmd.Flags().BoolVar(&opts.AllString, "all_string", false, "convert all values to strings")

	// Output format
	rootCmd.Flags().BoolVarP(&opts.ExportArray, "array", "a", false, "export as array (default: dict keyed by first column)")
	rootCmd.Flags().BoolVarP(&opts.ForceSheetName, "sheet", "s", false, "force sheet-name wrapping even for single sheet")
	rootCmd.Flags().BoolVar(&opts.Pretty, "pretty", false, "pretty-print JSON with tab indentation")
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		os.Exit(1)
	}
}

func run(files []string, opts Options) error {
	multi := len(files) > 1
	g, ctx := errgroup.WithContext(context.Background())
	g.SetLimit(2 * runtime.NumCPU())

	type result struct {
		file  string
		lines []string // output log lines
	}
	results := make([]result, len(files))

	for i, f := range files {
		g.Go(func() error {
			select {
			case <-ctx.Done():
				return ctx.Err()
			default:
			}

			sheets, err := LoadExcel(f, opts.HeaderRows)
			if err != nil {
				return fmt.Errorf("%s: %w", f, err)
			}

			target, targetIsDir := resolveTarget(multi, opts)
			var lines []string

			// stdout output must be serialized
			toStdout := target == ""
			if toStdout {
				stdoutMu.Lock()
			}

			switch opts.Format {
			case "json":
				path := target
				if targetIsDir {
					os.MkdirAll(target, 0755)
					path = filepath.Join(target, fileStem(f)+".json")
				}
				if err := writeJSON(sheets, opts, path); err != nil {
					if toStdout {
						stdoutMu.Unlock()
					}
					return fmt.Errorf("%s: %w", f, err)
				}
				if path != "" {
					lines = append(lines, path)
				}

			case "csv":
				path := target
				if targetIsDir {
					os.MkdirAll(target, 0755)
					path = filepath.Join(target, fileStem(f)+".csv")
				}
				if err := exportCSV(sheets, opts, path); err != nil {
					if toStdout {
						stdoutMu.Unlock()
					}
					return fmt.Errorf("%s: %w", f, err)
				}
				if path != "" {
					lines = append(lines, path)
				}

			default:
				if toStdout {
					stdoutMu.Unlock()
				}
				return fmt.Errorf("unknown format %q (must be json or csv)", opts.Format)
			}

			if toStdout {
				stdoutMu.Unlock()
			}

			results[i] = result{file: filepath.Base(f), lines: lines}
			return nil
		})
	}

	if err := g.Wait(); err != nil {
		return err
	}

	// print results in input order
	for _, r := range results {
		for _, l := range r.lines {
			fmt.Printf("→ %s\n", l)
		}
	}

	return nil
}

// resolveTarget returns the output path and whether it's a directory.
//   - no -o:                    "",    false (stdout)
//   - single input, file path:  path,  false
//   - directory or multi-input: dir,   true
func resolveTarget(multi bool, opts Options) (path string, isDir bool) {
	if opts.OutputPath == "" {
		return "", false
	}
	if isDirPath(opts.OutputPath) || multi {
		return opts.OutputPath, true
	}
	return opts.OutputPath, false
}

// writeJSON serialises sheets to JSON and writes to the target path.
func writeJSON(sheets []Sheet, opts Options, target string) error {
	content, err := ExportJSON(sheets, opts)
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

// expandInputs expands directories into their .xls/.xlsx file list.
// Non-directory args are kept as-is.
func expandInputs(args []string) []string {
	var out []string
	seen := make(map[string]bool)
	for _, a := range args {
		fi, err := os.Stat(a)
		if err != nil || !fi.IsDir() {
			// file or nonexistent — keep as-is (cobra already validated N>=1)
			out = append(out, a)
			continue
		}
		// directory — glob xls/xlsx
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

// isDirPath / fileStem defined in csv_exporter.go / types.go

func encodeContent(s, enc string) ([]byte, error) {
	switch enc {
	case "utf8-nobom":
		return []byte(s), nil
	default:
		return []byte(s), nil
	}
}
