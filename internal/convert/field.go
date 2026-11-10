package convert

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/actfuns/excel2json/internal/core"
	"github.com/actfuns/excel2json/internal/jsonpath"
	"github.com/actfuns/excel2json/internal/schema"
	"github.com/actfuns/excel2json/internal/typing"
)

// convertFieldSpec applies a parsed field annotation to a raw cell.
//
// Selectors only locate the nodes a rule applies to; they are not output
// selectors. Every rule rewrites the whole field in place and the field value
// is always returned as a whole.
func (e *Engine) convertFieldSpec(spec *schema.Spec, cell string, required *bool) (any, error) {
	if cell == "" {
		req := e.opts.RequireByDefault
		if required != nil {
			req = *required
		}
		if req {
			return nil, fmt.Errorf("field cannot be empty")
		}
		return schema.ZeroForSpec(spec), nil
	}

	hasSelector := false
	needJSON := false
	for _, rule := range spec.Rules {
		if rule.Selector != nil {
			hasSelector = true
			needJSON = true
		}
		if rule.Type != nil && typing.IsComposite(rule.Type) {
			needJSON = true
		}
	}

	var root any = cell
	jsonDecoded := false
	switch {
	case hasSelector:
		decoded, err := typing.DecodeJSONCell(strings.TrimSpace(cell))
		if err != nil {
			return nil, fmt.Errorf("expected JSON, got %q: %w", cell, err)
		}
		root = decoded
		jsonDecoded = true
	case needJSON:
		// A composite type without a selector can still be attempted from the
		// raw string. If the cell is not JSON, the normal conversion path will
		// report the usual error for that rule.
		if decoded, err := typing.DecodeJSONCell(strings.TrimSpace(cell)); err == nil {
			root = decoded
			jsonDecoded = true
		}
	}

	conditionalMatched := false
	for ri, rule := range spec.Rules {
		// The first matching conditional type wins; later conditional rules
		// for the same whole field are skipped.
		if rule.Conditional && conditionalMatched {
			continue
		}

		var nodes []jsonpath.Node
		if rule.Selector == nil {
			nodes = []jsonpath.Node{{Value: root}}
		} else {
			nodes = rule.Selector.Eval(root)
		}
		if len(nodes) == 0 {
			continue
		}

		type assignment struct {
			path  []jsonpath.Step
			value any
		}
		assignments := make([]assignment, 0, len(nodes))

		for _, node := range nodes {
			useJSON := jsonDecoded || rule.Selector != nil
			value, matched, err := e.convertAndCheck(node.Value, rule.Type, rule.Checks, useJSON, rule.Conditional)
			if err != nil {
				return nil, fmt.Errorf("rule %d: %w", ri+1, err)
			}
			if !matched {
				continue
			}
			assignments = append(assignments, assignment{path: node.Path, value: value})
			if rule.Conditional {
				conditionalMatched = true
			}
		}

		for _, a := range assignments {
			next, err := jsonpath.SetAtPath(root, a.path, a.value)
			if err != nil {
				return nil, fmt.Errorf("rule %d: %w", ri+1, err)
			}
			root = next
		}
	}

	return root, nil
}

// convertAndCheck converts one node and runs its checks. When conditional is
// true, a conversion failure means "this rule does not apply" and is reported
// as matched=false rather than as an error.
func (e *Engine) convertAndCheck(v any, typ *typing.Type, checks []*schema.Check, useJSON, conditional bool) (any, bool, error) {
	prepared, checks, err := e.preparePad(v, typ, checks)
	if err != nil {
		if conditional {
			return nil, false, nil
		}
		return nil, false, err
	}

	value, err := e.convertAlt(prepared, typ, useJSON)
	if err != nil {
		if conditional {
			return nil, false, nil
		}
		return nil, false, err
	}
	if err := e.runner.RunChecks(value, checks); err != nil {
		return nil, false, err
	}
	return value, true, nil
}

func (e *Engine) convertAlt(v any, t *typing.Type, useJSON bool) (any, error) {
	if t == nil {
		return v, nil
	}
	if useJSON {
		return e.conv.ConvertJSON(v, t)
	}
	if s, ok := v.(string); ok {
		return e.conv.ConvertString(s, t)
	}
	return e.conv.ConvertJSON(v, t)
}

func (e *Engine) preparePad(v any, typ *typing.Type, checks []*schema.Check) (any, []*schema.Check, error) {
	var pad *schema.Check
	rest := make([]*schema.Check, 0, len(checks))
	for _, ch := range checks {
		if ch.Name == "pad" {
			pad = ch
			continue
		}
		rest = append(rest, ch)
	}
	if pad == nil {
		return v, rest, nil
	}
	if typ == nil || typ.Kind != typing.KindArray {
		return nil, nil, fmt.Errorf("pad requires a fixed-size array type")
	}

	arr, ok := v.([]any)
	if !ok {
		return nil, nil, fmt.Errorf("pad expects an array, got %T", v)
	}
	if len(arr) > typ.Length {
		return nil, nil, fmt.Errorf("array length %d exceeds %d", len(arr), typ.Length)
	}
	if len(arr) == typ.Length {
		return arr, rest, nil
	}

	fill, err := e.padValue(typ.Elem, pad.Args)
	if err != nil {
		return nil, nil, err
	}
	out := make([]any, len(arr), typ.Length)
	copy(out, arr)
	for len(out) < typ.Length {
		out = append(out, fill)
	}
	return out, rest, nil
}

func (e *Engine) padValue(elem *typing.Type, args []string) (any, error) {
	if len(args) == 0 {
		return typing.ZeroValue(elem), nil
	}
	if typing.IsComposite(elem) {
		return nil, fmt.Errorf("pad value for composite element types is not supported")
	}
	return e.conv.ConvertString(schema.ArgLiteral(args[0]), elem)
}

// ---------------------------------------------------------------------------
// ref(...)
// ---------------------------------------------------------------------------

type refSetResult struct {
	set map[string]struct{}
	err error
}

func (e *Engine) CheckRef(v any, sheet, column string) error {
	set, err := e.refSet(sheet, column)
	if err != nil {
		return err
	}
	key := schema.CanonicalString(v)
	if _, found := set[key]; !found {
		return fmt.Errorf("value %q not found in %s.%s", key, sheet, column)
	}
	return nil
}

func (e *Engine) refSet(sheetName, column string) (map[string]struct{}, error) {
	cacheKey := sheetName + "\x00" + column
	if v, ok := e.refCache.Load(cacheKey); ok {
		r := v.(refSetResult)
		return r.set, r.err
	}
	set, err := e.buildRefSet(sheetName, column)
	e.refCache.Store(cacheKey, refSetResult{set: set, err: err})
	return set, err
}

// sheetPackageName returns the final output name (filename stem) used by
// ref(...) for a sheet. When no template is configured, the sheet name is the
// only available identity.
func (e *Engine) contextOutputName(ctx SheetContext) string {
	if ctx.OutputName != "" {
		return ctx.OutputName
	}
	if e.opts.NameTmpl == "" {
		return ctx.Sheet.Name
	}
	return core.SanitizeName(core.ExpandNameTemplate(e.opts.NameTmpl, e.opts.SourceStem, ctx.Sheet.Name))
}

func (e *Engine) findRefSheet(ref string) (*SheetContext, error) {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return nil, fmt.Errorf("ref: empty output name")
	}
	ref = filepath.Base(strings.TrimSuffix(ref, ".json"))

	var matches []*SheetContext
	for i := range e.sheets {
		name := e.contextOutputName(e.sheets[i])
		if strings.EqualFold(name, ref) || strings.EqualFold(core.SanitizeName(name), ref) {
			matches = append(matches, &e.sheets[i])
		}
	}
	switch len(matches) {
	case 1:
		return matches[0], nil
	case 0:
		return nil, fmt.Errorf("ref: output name %q not found", ref)
	default:
		return nil, fmt.Errorf("ref: output name %q is ambiguous across %d sheets", ref, len(matches))
	}
}

func (e *Engine) buildRefSet(sheetName, column string) (map[string]struct{}, error) {
	if len(e.sheets) == 0 {
		return nil, fmt.Errorf("ref(%s.%s) requires workbook context", sheetName, column)
	}

	ctx, err := e.findRefSheet(sheetName)
	if err != nil {
		return nil, err
	}
	sheet := &ctx.Sheet

	columnIndex := sheet.ColumnIndex(column)
	if columnIndex < 0 {
		return nil, fmt.Errorf("ref: column %q not found in sheet %q", column, sheet.Name)
	}

	colType := ""
	if columnIndex < len(sheet.Types) {
		colType = sheet.Types[columnIndex]
	}
	var spec *schema.Spec
	if colType != "" {
		spec, err = schema.Parse(colType, ctx.TypeDefs)
		if err != nil {
			return nil, fmt.Errorf("ref: sheet %q column %q type: %w", sheet.Name, column, err)
		}
	}

	set := make(map[string]struct{}, len(sheet.Rows))
	for ri, row := range sheet.Rows {
		cell := ""
		if columnIndex < len(row) {
			cell = strings.TrimSpace(row[columnIndex])
		}
		if cell == "" {
			continue
		}
		value, err := e.convertForRef(colType, spec, cell)
		if err != nil {
			return nil, fmt.Errorf("ref: sheet %q row %d column %q: %w", sheet.Name, ri+1, column, err)
		}
		set[schema.CanonicalString(value)] = struct{}{}
	}
	return set, nil
}

func (e *Engine) convertForRef(colType string, spec *schema.Spec, cell string) (any, error) {
	if colType == "" {
		return e.conv.LegacyConvert(cell)
	}
	if spec != nil && len(spec.Rules) == 1 && spec.Rules[0].Selector == nil {
		value, err := e.convertAlt(cell, spec.Rules[0].Type, false)
		if err == nil {
			return value, nil
		}
	}
	return cell, nil
}
