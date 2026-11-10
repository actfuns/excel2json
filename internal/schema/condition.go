package schema

import (
	"fmt"
	"strings"

	"github.com/actfuns/excel2json/internal/jsonpath"
	"github.com/actfuns/excel2json/internal/typing"
)

// CondOp is a condition comparison operator.
type CondOp int

const (
	OpEq CondOp = iota
	OpNe
	OpGt
	OpGte
	OpLt
	OpLte
)

func (o CondOp) String() string {
	switch o {
	case OpEq:
		return "="
	case OpNe:
		return "!="
	case OpGt:
		return ">"
	case OpGte:
		return ">="
	case OpLt:
		return "<"
	case OpLte:
		return "<="
	default:
		return "?"
	}
}

// Ref identifies the value used by a condition.
type Ref struct {
	Column string         // empty means the current cell
	Path   *jsonpath.Path // optional JSONPath applied to the referenced value
}

// CondKind identifies a condition expression node.
type CondKind int

const (
	CondCompare CondKind = iota
	CondType
	CondAnd
	CondOr
	CondNot
)

// Condition is one guard expression inside when(...). Combinators keep their
// operands in Subs.
type Condition struct {
	Kind     CondKind
	Ref      Ref
	Op       CondOp
	Literal  string
	TypeExpr *typing.Type
	Subs     []Condition
}

func parseConditions(text string) ([]Condition, error) {
	parts, err := splitTopLevel(text, ',')
	if err != nil {
		return nil, err
	}
	var out []Condition
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			return nil, fmt.Errorf("empty condition")
		}
		cond, err := parseCondition(part)
		if err != nil {
			return nil, err
		}
		out = append(out, cond)
	}
	return out, nil
}

func parseCondition(text string) (Condition, error) {
	text = strings.TrimSpace(text)
	lower := strings.ToLower(text)
	for _, comb := range []struct {
		prefix string
		kind   CondKind
	}{
		{"and(", CondAnd},
		{"or(", CondOr},
		{"not(", CondNot},
	} {
		if strings.HasPrefix(lower, comb.prefix) {
			open := strings.IndexByte(text, '(')
			close, err := findMatchingParen(text, open)
			if err != nil {
				return Condition{}, err
			}
			if strings.TrimSpace(text[close+1:]) != "" {
				return Condition{}, fmt.Errorf("unexpected trailing text after %s", comb.prefix)
			}
			subs, err := parseConditions(text[open+1 : close])
			if err != nil {
				return Condition{}, err
			}
			if len(subs) == 0 {
				return Condition{}, fmt.Errorf("%s requires at least one condition", strings.TrimSuffix(comb.prefix, "("))
			}
			return Condition{Kind: comb.kind, Subs: subs}, nil
		}
	}

	if strings.HasPrefix(lower, "typeof(") || strings.HasPrefix(lower, "type(") {
		open := strings.IndexByte(text, '(')
		close, err := findMatchingParen(text, open)
		if err != nil {
			return Condition{}, err
		}
		ref, err := parseRef(text[open+1 : close])
		if err != nil {
			return Condition{}, err
		}
		rest := strings.TrimSpace(text[close+1:])
		op, rhs, err := parseTypeOp(rest)
		if err != nil {
			return Condition{}, err
		}
		typeExpr, err := typing.Parse(rhs)
		if err != nil {
			return Condition{}, fmt.Errorf("invalid type %q: %w", rhs, err)
		}
		return Condition{Kind: CondType, Ref: ref, Op: op, TypeExpr: typeExpr}, nil
	}

	idx, op, opLen, err := findCondOp(text)
	if err != nil {
		return Condition{}, err
	}
	left := strings.TrimSpace(text[:idx])
	right := strings.TrimSpace(text[idx+opLen:])
	if left == "" || right == "" {
		return Condition{}, fmt.Errorf("invalid condition %q", text)
	}
	ref, err := parseRef(left)
	if err != nil {
		return Condition{}, err
	}
	return Condition{Kind: CondCompare, Ref: ref, Op: op, Literal: right}, nil
}

func parseRef(text string) (Ref, error) {
	text = strings.TrimSpace(text)
	if text == "" {
		return Ref{}, fmt.Errorf("empty condition reference")
	}
	if strings.HasPrefix(text, "$") {
		path, err := jsonpath.Parse(text)
		if err != nil {
			return Ref{}, err
		}
		return Ref{Path: path}, nil
	}
	if i := strings.IndexAny(text, ".["); i >= 0 {
		column := strings.TrimSpace(text[:i])
		if column == "" {
			return Ref{}, fmt.Errorf("empty column in %q", text)
		}
		path, err := jsonpath.Parse("$" + text[i:])
		if err != nil {
			return Ref{}, err
		}
		return Ref{Column: column, Path: path}, nil
	}
	return Ref{Column: text}, nil
}

func parseTypeOp(text string) (CondOp, string, error) {
	text = strings.TrimSpace(text)
	for _, candidate := range []struct {
		text string
		op   CondOp
	}{
		{"==", OpEq}, {"!=", OpNe}, {"=", OpEq},
	} {
		if strings.HasPrefix(text, candidate.text) {
			rhs := strings.TrimSpace(text[len(candidate.text):])
			if rhs == "" {
				return 0, "", fmt.Errorf("missing type after %q", candidate.text)
			}
			return candidate.op, rhs, nil
		}
	}
	return 0, "", fmt.Errorf("typeof condition requires =, == or !=")
}

func findCondOp(text string) (int, CondOp, int, error) {
	depth := 0
	var quote byte
	for i := 0; i < len(text); i++ {
		c := text[i]
		if quote != 0 {
			if c == '\\' {
				i++
				continue
			}
			if c == quote {
				quote = 0
			}
			continue
		}
		switch c {
		case '\'', '"':
			quote = c
		case '(', '[', '{':
			depth++
		case ')', ']', '}':
			if depth > 0 {
				depth--
			}
		default:
			if depth != 0 {
				continue
			}
			switch {
			case strings.HasPrefix(text[i:], "=="):
				return i, OpEq, 2, nil
			case strings.HasPrefix(text[i:], "!="):
				return i, OpNe, 2, nil
			case strings.HasPrefix(text[i:], ">="):
				return i, OpGte, 2, nil
			case strings.HasPrefix(text[i:], "<="):
				return i, OpLte, 2, nil
			case c == '=':
				return i, OpEq, 1, nil
			case c == '>':
				return i, OpGt, 1, nil
			case c == '<':
				return i, OpLt, 1, nil
			}
		}
	}
	return 0, 0, 0, fmt.Errorf("invalid condition %q: missing operator", text)
}

func findMatchingParen(s string, open int) (int, error) {
	if open < 0 || open >= len(s) || s[open] != '(' {
		return -1, fmt.Errorf("expected '('")
	}
	depth := 0
	var quote byte
	for i := open; i < len(s); i++ {
		c := s[i]
		if quote != 0 {
			if c == '\\' {
				i++
				continue
			}
			if c == quote {
				quote = 0
			}
			continue
		}
		switch c {
		case '\'', '"':
			quote = c
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				return i, nil
			}
		}
	}
	return -1, fmt.Errorf("unclosed '('")
}

// CompareValue evaluates a non-type condition against a value.
func CompareValue(v any, op CondOp, literal string) (bool, error) {
	switch op {
	case OpEq, OpNe:
		equal, err := equalLiteral(v, literal)
		if err != nil {
			return false, err
		}
		if op == OpNe {
			equal = !equal
		}
		return equal, nil
	}

	switch v.(type) {
	case string:
		cmp := strings.Compare(v.(string), ArgLiteral(literal))
		return cmpOp(cmp, op), nil
	default:
		cmp, err := compareNumericArg(v, literal)
		if err != nil {
			return false, err
		}
		return cmpOp(cmp, op), nil
	}
}

func cmpOp(cmp int, op CondOp) bool {
	switch op {
	case OpGt:
		return cmp > 0
	case OpGte:
		return cmp >= 0
	case OpLt:
		return cmp < 0
	case OpLte:
		return cmp <= 0
	}
	return false
}
