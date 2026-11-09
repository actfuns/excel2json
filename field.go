package main

import (
	"fmt"
	"path/filepath"
	"strings"
)

// convertFieldSpec applies a parsed field annotation to a raw cell.
//
// Selectors only locate the nodes a rule applies to; they are not output
// selectors. Every rule rewrites the whole field in place and the field value
// is always returned as a whole.
func (e *Exporter) convertFieldSpec(spec *fieldSpec, cell string, required *bool) (any, error) {
	if cell == "" {
		req := e.opts.Required
		if required != nil {
			req = *required
		}
		if req {
			return nil, fmt.Errorf("field cannot be empty")
		}
		return zeroForSpec(spec), nil
	}

	hasSelector := false
	needJSON := false
	for _, rule := range spec.rules {
		if rule.selector != nil {
			hasSelector = true
			needJSON = true
		}
		if rule.typ != nil && isCompositeType(rule.typ) {
			needJSON = true
		}
	}

	var root any = cell
	jsonDecoded := false
	switch {
	case hasSelector:
		decoded, err := decodeJSONCell(strings.TrimSpace(cell))
		if err != nil {
			return nil, fmt.Errorf("expected JSON, got %q: %w", cell, err)
		}
		root = decoded
		jsonDecoded = true
	case needJSON:
		// A composite type without a selector can still be attempted from the
		// raw string. If the cell is not JSON, the normal conversion path will
		// report the usual error for that rule.
		if decoded, err := decodeJSONCell(strings.TrimSpace(cell)); err == nil {
			root = decoded
			jsonDecoded = true
		}
	}

	conditionalMatched := false
	for ri, rule := range spec.rules {
		// The first matching conditional type wins; later conditional rules
		// for the same whole field are skipped.
		if rule.conditional && conditionalMatched {
			continue
		}

		var nodes []jpNode
		if rule.selector == nil {
			nodes = []jpNode{{value: root}}
		} else {
			nodes = rule.selector.eval(root)
		}
		if len(nodes) == 0 {
			continue
		}

		type assignment struct {
			path  []jpStep
			value any
		}
		assignments := make([]assignment, 0, len(nodes))

		for _, node := range nodes {
			useJSON := jsonDecoded || rule.selector != nil
			value, matched, err := e.convertAndCheck(node.value, rule.typ, rule.checks, useJSON, rule.conditional)
			if err != nil {
				return nil, fmt.Errorf("rule %d: %w", ri+1, err)
			}
			if !matched {
				continue
			}
			assignments = append(assignments, assignment{path: node.path, value: value})
			if rule.conditional {
				conditionalMatched = true
			}
		}

		for _, a := range assignments {
			next, err := setAtPath(root, a.path, a.value)
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
func (e *Exporter) convertAndCheck(v any, typ *goType, checks []*checkExpr, useJSON, conditional bool) (any, bool, error) {
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
	if err := e.runChecks(value, checks); err != nil {
		return nil, false, err
	}
	return value, true, nil
}

func (e *Exporter) convertAlt(v any, t *goType, useJSON bool) (any, error) {
	if t == nil {
		return v, nil
	}
	if useJSON {
		return e.convertJSON(v, t)
	}
	if s, ok := v.(string); ok {
		return e.convertString(s, t)
	}
	return e.convertJSON(v, t)
}

func (e *Exporter) preparePad(v any, typ *goType, checks []*checkExpr) (any, []*checkExpr, error) {
	var pad *checkExpr
	rest := make([]*checkExpr, 0, len(checks))
	for _, ch := range checks {
		if ch.name == "pad" {
			pad = ch
			continue
		}
		rest = append(rest, ch)
	}
	if pad == nil {
		return v, rest, nil
	}
	if typ == nil || typ.kind != kindArray {
		return nil, nil, fmt.Errorf("pad requires a fixed-size array type")
	}

	arr, ok := v.([]any)
	if !ok {
		return nil, nil, fmt.Errorf("pad expects an array, got %T", v)
	}
	if len(arr) > typ.length {
		return nil, nil, fmt.Errorf("array length %d exceeds %d", len(arr), typ.length)
	}
	if len(arr) == typ.length {
		return arr, rest, nil
	}

	fill, err := e.padValue(typ.elem, pad.args)
	if err != nil {
		return nil, nil, err
	}
	out := make([]any, len(arr), typ.length)
	copy(out, arr)
	for len(out) < typ.length {
		out = append(out, fill)
	}
	return out, rest, nil
}

func (e *Exporter) padValue(elem *goType, args []string) (any, error) {
	if len(args) == 0 {
		return zeroValue(elem), nil
	}
	if isCompositeType(elem) {
		return nil, fmt.Errorf("pad value for composite element types is not supported")
	}
	return e.convertString(argLiteral(args[0]), elem)
}

func zeroForSpec(spec *fieldSpec) any {
	for _, rule := range spec.rules {
		if rule.typ == nil {
			continue
		}
		if rule.typ.kind == kindArray && hasPad(rule.checks) {
			out := make([]any, 0, rule.typ.length)
			fill := zeroValue(rule.typ.elem)
			for i := 0; i < rule.typ.length; i++ {
				out = append(out, fill)
			}
			return out
		}
		return zeroValue(rule.typ)
	}
	return ""
}

func hasPad(checks []*checkExpr) bool {
	for _, ch := range checks {
		if ch.name == "pad" {
			return true
		}
	}
	return false
}

func isCompositeType(t *goType) bool {
	return t.kind == kindSlice || t.kind == kindArray || t.kind == kindMap
}

// ---------------------------------------------------------------------------
// ref(...)
// ---------------------------------------------------------------------------

type refSetResult struct {
	set map[string]struct{}
	err error
}

func (e *Exporter) checkRef(v any, arg string) error {
	arg = strings.TrimSpace(arg)
	dot := strings.LastIndexByte(arg, '.')
	if dot <= 0 || dot == len(arg)-1 {
		return fmt.Errorf("ref expects package.column, got %q", arg)
	}
	sheet, column := arg[:dot], arg[dot+1:]
	set, err := e.refSet(sheet, column)
	if err != nil {
		return err
	}
	key := canonicalString(v)
	if _, found := set[key]; !found {
		return fmt.Errorf("value %q not found in %s.%s", key, sheet, column)
	}
	return nil
}

func (e *Exporter) refSet(sheetName, column string) (map[string]struct{}, error) {
	cacheKey := sheetName + "\x00" + column
	if v, ok := e.refCache.Load(cacheKey); ok {
		r := v.(refSetResult)
		return r.set, r.err
	}
	set, err := e.buildRefSet(sheetName, column)
	e.refCache.Store(cacheKey, refSetResult{set: set, err: err})
	return set, err
}

// sheetPackageName returns the logical output name (JSON filename stem) used
// by ref(...) for a sheet. With --merge there is only one output file, so the
// sheet name is used instead.
func (e *Exporter) sheetPackageName(sheetName string) string {
	if e.opts.Merge || e.opts.NameTmpl == "" {
		return sheetName
	}
	return formatName(e.opts.NameTmpl, e.opts.SourceName, sheetName)
}

func (e *Exporter) findRefSheet(ref string) (*Sheet, error) {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return nil, fmt.Errorf("ref: empty sheet/package name")
	}
	ref = filepath.Base(strings.TrimSuffix(ref, ".json"))

	var matches []*Sheet
	for i := range e.sheets {
		pkg := e.sheetPackageName(e.sheets[i].Name)
		if strings.EqualFold(pkg, ref) || strings.EqualFold(sanitizeName(pkg), ref) {
			matches = append(matches, &e.sheets[i])
		}
	}
	switch len(matches) {
	case 1:
		return matches[0], nil
	case 0:
		// Fall back to the raw sheet name for configurations that either use
		// no naming template or reference the sheet directly.
		for i := range e.sheets {
			if strings.EqualFold(e.sheets[i].Name, ref) {
				return &e.sheets[i], nil
			}
		}
		return nil, fmt.Errorf("ref: package %q not found", ref)
	default:
		return nil, fmt.Errorf("ref: package %q is ambiguous across %d sheets", ref, len(matches))
	}
}

func (e *Exporter) buildRefSet(sheetName, column string) (map[string]struct{}, error) {
	if len(e.sheets) == 0 {
		return nil, fmt.Errorf("ref(%s.%s) requires workbook context", sheetName, column)
	}

	sheet, err := e.findRefSheet(sheetName)
	if err != nil {
		return nil, err
	}

	columnIndex := -1
	for i, header := range sheet.Headers {
		if header == column || strings.EqualFold(header, column) {
			columnIndex = i
			break
		}
	}
	if columnIndex < 0 {
		return nil, fmt.Errorf("ref: column %q not found in sheet %q", column, sheet.Name)
	}

	colType := ""
	if columnIndex < len(sheet.Types) {
		colType = sheet.Types[columnIndex]
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
		value, err := e.convertForRef(colType, cell)
		if err != nil {
			return nil, fmt.Errorf("ref: sheet %q row %d column %q: %w", sheet.Name, ri+1, column, err)
		}
		set[canonicalString(value)] = struct{}{}
	}
	return set, nil
}

func (e *Exporter) convertForRef(colType, cell string) (any, error) {
	if colType == "" {
		return e.legacyConvert(cell)
	}
	spec, err := e.parseFieldSpecCached(colType)
	if err != nil {
		return nil, err
	}
	if len(spec.rules) == 1 && spec.rules[0].selector == nil {
		value, err := e.convertAlt(cell, spec.rules[0].typ, false)
		if err == nil {
			return value, nil
		}
	}
	return cell, nil
}
