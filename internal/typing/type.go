package typing

import (
	"encoding/json"
	"fmt"
	"io"
	"math"
	"strconv"
	"strings"
	"unicode"
)

// Kind identifies a parsed Go type expression.
type Kind int

const (
	KindAny Kind = iota
	KindBool
	KindString
	KindInt
	KindUint
	KindFloat
	KindDate
	KindSlice
	KindArray
	KindMap
	KindNull
)

// Type is a parsed subset of Go's type syntax.
// Supported forms:
//
//	int int8 int16 int32 int64 uint uint8 uint16 uint32 uint64 uintptr
//	float32 float64 bool string any
//	[]T  [N]T  map[K]V
//	额外支持: date（日期）, float（float64 别名）, object（map[string]any 别名）, list（[]any 别名）, null
type Type struct {
	Kind   Kind
	Name   string // original name for scalar types
	Bits   int    // bit width for int/uint/float
	Length int    // array length
	Key    *Type
	Elem   *Type
}

// Converter converts raw cell strings and decoded JSON values into typed values.
type Converter struct {
	DateFormat    string
	ParseCellJSON bool
	StringifyAll  bool
}

func (t *Type) String() string {
	if t == nil {
		return "<nil>"
	}
	var s string
	switch t.Kind {
	case KindSlice:
		s = "[]" + t.Elem.String()
	case KindArray:
		s = "[" + strconv.Itoa(t.Length) + "]" + t.Elem.String()
	case KindMap:
		s = "map[" + t.Key.String() + "]" + t.Elem.String()
	default:
		s = t.Name
	}
	return s
}

// typeParser is a small recursive-descent parser for Go type expressions.
type typeParser struct {
	s   string
	pos int
}

func Parse(s string) (*Type, error) {
	var b strings.Builder
	for _, r := range s {
		if !unicode.IsSpace(r) {
			b.WriteRune(r)
		}
	}
	norm := b.String()

	if norm == "" {
		return nil, fmt.Errorf("invalid type %q: empty type", s)
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
	return t, nil
}

func (p *typeParser) parse() (*Type, error) {
	return p.parsePrimary()
}

func (p *typeParser) parsePrimary() (*Type, error) {
	rest := p.s[p.pos:]

	switch {
	case strings.HasPrefix(rest, "[]"):
		p.pos += 2
		elem, err := p.parse()
		if err != nil {
			return nil, err
		}
		return &Type{Kind: KindSlice, Elem: elem}, nil

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
		return &Type{Kind: KindArray, Length: n, Elem: elem}, nil

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
		return &Type{Kind: KindMap, Key: key, Elem: elem}, nil

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

func basicType(name string) (*Type, error) {
	canonical := strings.ToLower(name)
	switch canonical {
	case "any":
		return &Type{Kind: KindAny, Name: canonical}, nil
	case "bool":
		return &Type{Kind: KindBool, Name: canonical}, nil
	case "string":
		return &Type{Kind: KindString, Name: canonical}, nil
	case "date":
		return &Type{Kind: KindDate, Name: canonical}, nil
	case "null":
		return &Type{Kind: KindNull, Name: canonical}, nil
	case "object":
		return &Type{
			Kind: KindMap,
			Name: canonical,
			Key:  &Type{Kind: KindString, Name: "string"},
			Elem: &Type{Kind: KindAny, Name: "any"},
		}, nil
	case "list":
		return &Type{
			Kind: KindSlice,
			Name: canonical,
			Elem: &Type{Kind: KindAny, Name: "any"},
		}, nil
	case "int":
		return &Type{Kind: KindInt, Name: canonical, Bits: 64}, nil
	case "int8":
		return &Type{Kind: KindInt, Name: canonical, Bits: 8}, nil
	case "int16":
		return &Type{Kind: KindInt, Name: canonical, Bits: 16}, nil
	case "int32", "rune":
		return &Type{Kind: KindInt, Name: canonical, Bits: 32}, nil
	case "int64":
		return &Type{Kind: KindInt, Name: canonical, Bits: 64}, nil
	case "uint":
		return &Type{Kind: KindUint, Name: canonical, Bits: 64}, nil
	case "uint8", "byte":
		return &Type{Kind: KindUint, Name: canonical, Bits: 8}, nil
	case "uint16":
		return &Type{Kind: KindUint, Name: canonical, Bits: 16}, nil
	case "uint32":
		return &Type{Kind: KindUint, Name: canonical, Bits: 32}, nil
	case "uint64", "uintptr":
		return &Type{Kind: KindUint, Name: canonical, Bits: 64}, nil
	case "float", "float64":
		return &Type{Kind: KindFloat, Name: canonical, Bits: 64}, nil
	case "float32":
		return &Type{Kind: KindFloat, Name: canonical, Bits: 32}, nil
	default:
		return nil, fmt.Errorf("unsupported type %q", name)
	}
}

// IsComposite reports whether t is a slice, array or map.
func IsComposite(t *Type) bool {
	return t.Kind == KindSlice || t.Kind == KindArray || t.Kind == KindMap
}

// ZeroValue returns the value emitted for an empty typed cell.
func ZeroValue(t *Type) any {
	switch t.Kind {
	case KindBool:
		return false
	case KindString, KindDate:
		return ""
	case KindInt:
		return int64(0)
	case KindUint:
		return uint64(0)
	case KindFloat:
		return float64(0)
	default:
		return nil
	}
}

// convertString converts a raw cell string to the target type.
func (c *Converter) ConvertString(cell string, t *Type) (any, error) {
	switch t.Kind {
	case KindString:
		return cell, nil

	case KindNull:
		if strings.TrimSpace(cell) == "null" {
			return nil, nil
		}
		return nil, fmt.Errorf("expected null, got %q", cell)

	case KindBool:
		v, err := strconv.ParseBool(strings.ToLower(strings.TrimSpace(cell)))
		if err != nil {
			return nil, fmt.Errorf("expected %s, got %q", t, cell)
		}
		return v, nil

	case KindInt:
		v, err := strconv.ParseInt(strings.TrimSpace(cell), 10, t.Bits)
		if err != nil {
			return nil, fmt.Errorf("expected %s, got %q", t, cell)
		}
		return v, nil

	case KindUint:
		v, err := strconv.ParseUint(strings.TrimSpace(cell), 10, t.Bits)
		if err != nil {
			return nil, fmt.Errorf("expected %s, got %q", t, cell)
		}
		return v, nil

	case KindFloat:
		v, err := strconv.ParseFloat(strings.TrimSpace(cell), t.Bits)
		if err != nil || math.IsNaN(v) || math.IsInf(v, 0) {
			return nil, fmt.Errorf("expected %s, got %q", t, cell)
		}
		return v, nil

	case KindDate:
		return c.ConvertDate(cell)

	case KindAny:
		return c.LegacyConvert(cell)

	default:
		trimmed := strings.TrimSpace(cell)
		parsed, err := DecodeJSONCell(trimmed)
		if err != nil {
			return nil, fmt.Errorf("expected JSON for %s, got %q", t, cell)
		}
		return c.ConvertJSON(parsed, t)
	}
}

// DecodeJSONCell decodes exactly one JSON value, preserving number literals
// as json.Number so nested int64/uint64 validation stays exact.
func DecodeJSONCell(s string) (any, error) {
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
func (c *Converter) ConvertJSON(v any, t *Type) (any, error) {
	switch t.Kind {
	case KindAny:
		return v, nil

	case KindNull:
		if v == nil {
			return nil, nil
		}
		return nil, fmt.Errorf("expected null, got %T", v)

	case KindString:
		s, ok := v.(string)
		if !ok {
			return nil, fmt.Errorf("expected string, got %T", v)
		}
		return s, nil

	case KindBool:
		b, ok := v.(bool)
		if !ok {
			return nil, fmt.Errorf("expected bool, got %T", v)
		}
		return b, nil

	case KindInt:
		if n, ok := v.(json.Number); ok {
			if iv, err := strconv.ParseInt(n.String(), 10, t.Bits); err == nil {
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

	case KindUint:
		if n, ok := v.(json.Number); ok {
			if uv, err := strconv.ParseUint(n.String(), 10, t.Bits); err == nil {
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

	case KindFloat:
		f, ok := jsonNumber(v)
		if !ok {
			return nil, fmt.Errorf("expected number, got %T", v)
		}
		if t.Bits == 32 && math.Abs(f) > math.MaxFloat32 {
			return nil, fmt.Errorf("value %v overflows %s", f, t)
		}
		return f, nil

	case KindDate:
		switch d := v.(type) {
		case string:
			return c.ConvertDate(d)
		default:
			f, ok := jsonNumber(v)
			if !ok {
				return nil, fmt.Errorf("expected date string or serial number, got %T", v)
			}
			return c.ConvertDate(strconv.FormatFloat(f, 'f', -1, 64))
		}

	case KindSlice:
		arr, ok := v.([]any)
		if !ok {
			return nil, fmt.Errorf("expected array, got %T", v)
		}
		out := make([]any, len(arr))
		for i, item := range arr {
			cv, err := c.ConvertJSON(item, t.Elem)
			if err != nil {
				return nil, fmt.Errorf("index %d: %w", i, err)
			}
			out[i] = cv
		}
		return out, nil

	case KindArray:
		arr, ok := v.([]any)
		if !ok {
			return nil, fmt.Errorf("expected array, got %T", v)
		}
		if len(arr) != t.Length {
			return nil, fmt.Errorf("expected array of length %d, got %d", t.Length, len(arr))
		}
		out := make([]any, len(arr))
		for i, item := range arr {
			cv, err := c.ConvertJSON(item, t.Elem)
			if err != nil {
				return nil, fmt.Errorf("index %d: %w", i, err)
			}
			out[i] = cv
		}
		return out, nil

	case KindMap:
		obj, ok := v.(map[string]any)
		if !ok {
			return nil, fmt.Errorf("expected object, got %T", v)
		}
		out := make(map[string]any, len(obj))
		for k, item := range obj {
			key, err := c.ConvertString(k, t.Key)
			if err != nil {
				return nil, fmt.Errorf("key %q: %w", k, err)
			}
			cv, err := c.ConvertJSON(item, t.Elem)
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

func intValue(f float64, t *Type) (int64, error) {
	min := -math.Pow(2, float64(t.Bits-1))
	maxExclusive := math.Pow(2, float64(t.Bits-1))
	if f < min || f >= maxExclusive {
		return 0, fmt.Errorf("value %v out of range for %s", f, t)
	}
	return int64(f), nil
}

func uintValue(f float64, t *Type) (uint64, error) {
	maxExclusive := math.Pow(2, float64(t.Bits))
	if f < 0 || f >= maxExclusive {
		return 0, fmt.Errorf("value %v out of range for %s", f, t)
	}
	return uint64(f), nil
}
