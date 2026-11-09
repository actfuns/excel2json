package main

import (
	"encoding/json"
	"fmt"
	"io"
	"math"
	"strconv"
	"strings"
	"unicode"
)

// typeKind identifies a parsed Go type expression.
type typeKind int

const (
	kindAny typeKind = iota
	kindBool
	kindString
	kindInt
	kindUint
	kindFloat
	kindDate
	kindSlice
	kindArray
	kindMap
)

// goType is a parsed subset of Go's type syntax.
// Supported forms:
//
//	int int8 int16 int32 int64 uint uint8 uint16 uint32 uint64 uintptr
//	float32 float64 bool string any
//	[]T  [N]T  map[K]V
//	额外支持: date（日期）, float（float64 别名）, object（map[string]any 别名）, list（[]any 别名）
//
// parseType 仍支持末尾的 !/? 标记（直接调用时作为空值/必填语义），
// 但在字段注解中 !/? 由 annotation 层处理为类型严格/条件匹配。
type goType struct {
	kind     typeKind
	name     string // original name for scalar types
	bits     int    // bit width for int/uint/float
	length   int    // array length
	key      *goType
	elem     *goType
	nullable *bool // nil = use global default; true = ?, false = !
}

func (t *goType) String() string {
	if t == nil {
		return "<nil>"
	}
	var s string
	switch t.kind {
	case kindSlice:
		s = "[]" + t.elem.String()
	case kindArray:
		s = "[" + strconv.Itoa(t.length) + "]" + t.elem.String()
	case kindMap:
		s = "map[" + t.key.String() + "]" + t.elem.String()
	default:
		s = t.name
	}
	if t.nullable != nil {
		if *t.nullable {
			s += "?"
		} else {
			s += "!"
		}
	}
	return s
}

// typeParser is a small recursive-descent parser for Go type expressions.
type typeParser struct {
	s   string
	pos int
}

func parseType(s string) (*goType, error) {
	var b strings.Builder
	for _, r := range s {
		if !unicode.IsSpace(r) {
			b.WriteRune(r)
		}
	}
	norm := b.String()

	// Nullability marker is only allowed at the very end of the type.
	var nullable *bool
	if strings.HasSuffix(norm, "!") {
		v := false
		nullable = &v
		norm = strings.TrimSuffix(norm, "!")
	} else if strings.HasSuffix(norm, "?") {
		v := true
		nullable = &v
		norm = strings.TrimSuffix(norm, "?")
	}
	if norm == "" {
		return nil, fmt.Errorf("invalid type %q: missing type before ?/!", s)
	}

	norm = strings.ReplaceAll(norm, "interface{}", "any")

	p := &typeParser{s: norm}
	t, err := p.parse()
	if err != nil {
		return nil, err
	}
	if p.pos != len(p.s) {
		return nil, fmt.Errorf("invalid type %q: unexpected %q", s, p.s[p.pos:])
	}
	t.nullable = nullable
	return t, nil
}

func (p *typeParser) parse() (*goType, error) {
	return p.parsePrimary()
}

func (p *typeParser) parsePrimary() (*goType, error) {
	rest := p.s[p.pos:]

	switch {
	case strings.HasPrefix(rest, "[]"):
		p.pos += 2
		elem, err := p.parse()
		if err != nil {
			return nil, err
		}
		return &goType{kind: kindSlice, elem: elem}, nil

	case strings.HasPrefix(rest, "["):
		end := strings.IndexByte(rest, ']')
		if end < 0 {
			return nil, fmt.Errorf("invalid array type: missing ']'")
		}
		lenStr := rest[1:end]
		n, err := strconv.Atoi(lenStr)
		if err != nil || n < 0 {
			return nil, fmt.Errorf("invalid array length %q", lenStr)
		}
		p.pos += end + 1
		elem, err := p.parse()
		if err != nil {
			return nil, err
		}
		return &goType{kind: kindArray, length: n, elem: elem}, nil

	case strings.HasPrefix(rest, "map["):
		p.pos += len("map[")
		key, err := p.parse()
		if err != nil {
			return nil, err
		}
		if err := p.expect(']'); err != nil {
			return nil, err
		}
		elem, err := p.parse()
		if err != nil {
			return nil, err
		}
		return &goType{kind: kindMap, key: key, elem: elem}, nil

	default:
		start := p.pos
		for p.pos < len(p.s) && isTypeIdentChar(p.s[p.pos]) {
			p.pos++
		}
		if start == p.pos {
			if p.pos >= len(p.s) {
				return nil, fmt.Errorf("unexpected end of type")
			}
			return nil, fmt.Errorf("unexpected character %q in type", p.s[p.pos])
		}
		return basicType(p.s[start:p.pos])
	}
}

func (p *typeParser) expect(c byte) error {
	if p.pos >= len(p.s) || p.s[p.pos] != c {
		return fmt.Errorf("expected %q in type", c)
	}
	p.pos++
	return nil
}

func isTypeIdentChar(c byte) bool {
	return c == '_' ||
		(c >= 'a' && c <= 'z') ||
		(c >= 'A' && c <= 'Z') ||
		(c >= '0' && c <= '9')
}

func basicType(name string) (*goType, error) {
	canonical := strings.ToLower(name)
	switch canonical {
	case "any":
		return &goType{kind: kindAny, name: canonical}, nil
	case "bool":
		return &goType{kind: kindBool, name: canonical}, nil
	case "string":
		return &goType{kind: kindString, name: canonical}, nil
	case "date":
		return &goType{kind: kindDate, name: canonical}, nil
	case "object":
		return &goType{
			kind: kindMap,
			name: canonical,
			key:  &goType{kind: kindString, name: "string"},
			elem: &goType{kind: kindAny, name: "any"},
		}, nil
	case "list":
		return &goType{
			kind: kindSlice,
			name: canonical,
			elem: &goType{kind: kindAny, name: "any"},
		}, nil
	case "int":
		return &goType{kind: kindInt, name: canonical, bits: 64}, nil
	case "int8":
		return &goType{kind: kindInt, name: canonical, bits: 8}, nil
	case "int16":
		return &goType{kind: kindInt, name: canonical, bits: 16}, nil
	case "int32", "rune":
		return &goType{kind: kindInt, name: canonical, bits: 32}, nil
	case "int64":
		return &goType{kind: kindInt, name: canonical, bits: 64}, nil
	case "uint":
		return &goType{kind: kindUint, name: canonical, bits: 64}, nil
	case "uint8", "byte":
		return &goType{kind: kindUint, name: canonical, bits: 8}, nil
	case "uint16":
		return &goType{kind: kindUint, name: canonical, bits: 16}, nil
	case "uint32":
		return &goType{kind: kindUint, name: canonical, bits: 32}, nil
	case "uint64", "uintptr":
		return &goType{kind: kindUint, name: canonical, bits: 64}, nil
	case "float", "float64":
		return &goType{kind: kindFloat, name: canonical, bits: 64}, nil
	case "float32":
		return &goType{kind: kindFloat, name: canonical, bits: 32}, nil
	default:
		return nil, fmt.Errorf("unsupported type %q", name)
	}
}

// zeroValue returns the value emitted for an empty typed cell.
func zeroValue(t *goType) any {
	switch t.kind {
	case kindBool:
		return false
	case kindString, kindDate:
		return ""
	case kindInt:
		return int64(0)
	case kindUint:
		return uint64(0)
	case kindFloat:
		return float64(0)
	default:
		return nil
	}
}

type cachedType struct {
	t   *goType
	err error
}

// parseTypeCached parses a type once per Exporter.
func (e *Exporter) parseTypeCached(colType string) (*goType, error) {
	if v, ok := e.typeCache.Load(colType); ok {
		c := v.(cachedType)
		return c.t, c.err
	}
	t, err := parseType(colType)
	e.typeCache.Store(colType, cachedType{t: t, err: err})
	return t, err
}

// convertTyped parses colType and converts cell according to it.
func (e *Exporter) convertTyped(colType, cell string) (any, error) {
	t, err := e.parseTypeCached(colType)
	if err != nil {
		return nil, err
	}
	if cell == "" {
		if e.isRequired(t) {
			return nil, fmt.Errorf("field cannot be empty (type %s)", t)
		}
		return zeroValue(t), nil
	}
	return e.convertString(cell, t)
}

// isRequired reports whether an empty cell is an error for this type.
// Explicit ?/! wins; otherwise --required decides.
func (e *Exporter) isRequired(t *goType) bool {
	if t.nullable != nil {
		return !*t.nullable
	}
	return e.opts.Required
}

// convertString converts a raw cell string to the target type.
func (e *Exporter) convertString(cell string, t *goType) (any, error) {
	switch t.kind {
	case kindString:
		return cell, nil

	case kindBool:
		v, err := strconv.ParseBool(strings.ToLower(strings.TrimSpace(cell)))
		if err != nil {
			return nil, fmt.Errorf("expected %s, got %q", t, cell)
		}
		return v, nil

	case kindInt:
		v, err := strconv.ParseInt(strings.TrimSpace(cell), 10, t.bits)
		if err != nil {
			return nil, fmt.Errorf("expected %s, got %q", t, cell)
		}
		return v, nil

	case kindUint:
		v, err := strconv.ParseUint(strings.TrimSpace(cell), 10, t.bits)
		if err != nil {
			return nil, fmt.Errorf("expected %s, got %q", t, cell)
		}
		return v, nil

	case kindFloat:
		v, err := strconv.ParseFloat(strings.TrimSpace(cell), t.bits)
		if err != nil || math.IsNaN(v) || math.IsInf(v, 0) {
			return nil, fmt.Errorf("expected %s, got %q", t, cell)
		}
		return v, nil

	case kindDate:
		return e.convertDate(cell)

	case kindAny:
		return e.legacyConvert(cell)

	default:
		trimmed := strings.TrimSpace(cell)
		parsed, err := decodeJSONCell(trimmed)
		if err != nil {
			return nil, fmt.Errorf("expected JSON for %s, got %q", t, cell)
		}
		return e.convertJSON(parsed, t)
	}
}

// decodeJSONCell decodes exactly one JSON value, preserving number literals
// as json.Number so nested int64/uint64 validation stays exact.
func decodeJSONCell(s string) (any, error) {
	dec := json.NewDecoder(strings.NewReader(s))
	dec.UseNumber()

	var v any
	if err := dec.Decode(&v); err != nil {
		return nil, err
	}

	var extra any
	if err := dec.Decode(&extra); err != io.EOF {
		if err == nil {
			return nil, fmt.Errorf("unexpected trailing JSON value")
		}
		return nil, err
	}
	return v, nil
}

// convertJSON converts a decoded JSON value to the target type.
func (e *Exporter) convertJSON(v any, t *goType) (any, error) {
	switch t.kind {
	case kindAny:
		return v, nil

	case kindString:
		s, ok := v.(string)
		if !ok {
			return nil, fmt.Errorf("expected string, got %T", v)
		}
		return s, nil

	case kindBool:
		b, ok := v.(bool)
		if !ok {
			return nil, fmt.Errorf("expected bool, got %T", v)
		}
		return b, nil

	case kindInt:
		if n, ok := v.(json.Number); ok {
			if iv, err := strconv.ParseInt(n.String(), 10, t.bits); err == nil {
				return iv, nil
			}
			f, err := n.Float64()
			if err != nil {
				return nil, fmt.Errorf("expected integer, got %v", n)
			}
			if f != math.Trunc(f) {
				return nil, fmt.Errorf("expected integer, got %v", n)
			}
			return intValue(f, t)
		}
		f, ok := jsonNumber(v)
		if !ok {
			return nil, fmt.Errorf("expected integer, got %T", v)
		}
		if f != math.Trunc(f) {
			return nil, fmt.Errorf("expected integer, got %v", v)
		}
		return intValue(f, t)

	case kindUint:
		if n, ok := v.(json.Number); ok {
			if uv, err := strconv.ParseUint(n.String(), 10, t.bits); err == nil {
				return uv, nil
			}
			f, err := n.Float64()
			if err != nil {
				return nil, fmt.Errorf("expected unsigned integer, got %v", n)
			}
			if f != math.Trunc(f) || f < 0 {
				return nil, fmt.Errorf("expected unsigned integer, got %v", n)
			}
			return uintValue(f, t)
		}
		f, ok := jsonNumber(v)
		if !ok {
			return nil, fmt.Errorf("expected unsigned integer, got %T", v)
		}
		if f != math.Trunc(f) || f < 0 {
			return nil, fmt.Errorf("expected unsigned integer, got %v", v)
		}
		return uintValue(f, t)

	case kindFloat:
		f, ok := jsonNumber(v)
		if !ok {
			return nil, fmt.Errorf("expected number, got %T", v)
		}
		if t.bits == 32 && math.Abs(f) > math.MaxFloat32 {
			return nil, fmt.Errorf("value %v overflows %s", f, t)
		}
		return f, nil

	case kindDate:
		switch d := v.(type) {
		case string:
			return e.convertDate(d)
		default:
			f, ok := jsonNumber(v)
			if !ok {
				return nil, fmt.Errorf("expected date string or serial number, got %T", v)
			}
			return e.convertDate(strconv.FormatFloat(f, 'f', -1, 64))
		}

	case kindSlice:
		arr, ok := v.([]any)
		if !ok {
			return nil, fmt.Errorf("expected array, got %T", v)
		}
		out := make([]any, len(arr))
		for i, item := range arr {
			cv, err := e.convertJSON(item, t.elem)
			if err != nil {
				return nil, fmt.Errorf("index %d: %w", i, err)
			}
			out[i] = cv
		}
		return out, nil

	case kindArray:
		arr, ok := v.([]any)
		if !ok {
			return nil, fmt.Errorf("expected array, got %T", v)
		}
		if len(arr) != t.length {
			return nil, fmt.Errorf("expected array of length %d, got %d", t.length, len(arr))
		}
		out := make([]any, len(arr))
		for i, item := range arr {
			cv, err := e.convertJSON(item, t.elem)
			if err != nil {
				return nil, fmt.Errorf("index %d: %w", i, err)
			}
			out[i] = cv
		}
		return out, nil

	case kindMap:
		obj, ok := v.(map[string]any)
		if !ok {
			return nil, fmt.Errorf("expected object, got %T", v)
		}
		out := make(map[string]any, len(obj))
		for k, item := range obj {
			key, err := e.convertString(k, t.key)
			if err != nil {
				return nil, fmt.Errorf("key %q: %w", k, err)
			}
			cv, err := e.convertJSON(item, t.elem)
			if err != nil {
				return nil, fmt.Errorf("key %q: %w", k, err)
			}
			out[fmt.Sprint(key)] = cv
		}
		return out, nil
	}

	return nil, fmt.Errorf("unsupported type %s", t)
}

// jsonNumber extracts a float64 from a decoded JSON number.
func jsonNumber(v any) (float64, bool) {
	switch n := v.(type) {
	case float64:
		return n, true
	case float32:
		return float64(n), true
	case int:
		return float64(n), true
	case int8:
		return float64(n), true
	case int16:
		return float64(n), true
	case int32:
		return float64(n), true
	case int64:
		return float64(n), true
	case uint:
		return float64(n), true
	case uint8:
		return float64(n), true
	case uint16:
		return float64(n), true
	case uint32:
		return float64(n), true
	case uint64:
		return float64(n), true
	case json.Number:
		f, err := n.Float64()
		return f, err == nil
	default:
		return 0, false
	}
}

func intValue(f float64, t *goType) (int64, error) {
	min := -math.Pow(2, float64(t.bits-1))
	maxExclusive := math.Pow(2, float64(t.bits-1))
	if f < min || f >= maxExclusive {
		return 0, fmt.Errorf("value %v out of range for %s", f, t)
	}
	return int64(f), nil
}

func uintValue(f float64, t *goType) (uint64, error) {
	maxExclusive := math.Pow(2, float64(t.bits))
	if f < 0 || f >= maxExclusive {
		return 0, fmt.Errorf("value %v out of range for %s", f, t)
	}
	return uint64(f), nil
}
