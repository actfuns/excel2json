// Package schema implements the field annotation language: rules, JSONPath
// selectors, validators, combinators and named type definitions.
package schema

import (
	"fmt"
	"strings"

	"github.com/actfuns/excel2json/internal/jsonpath"
	"github.com/actfuns/excel2json/internal/typing"
)

// Spec is a parsed field annotation.
//
//	annotation = rule { ";" rule }
//	rule       = [ selector ":" ] [ type [ "!" | "?" ] [ ":" checks ] ]
//	checks     = check { "," check }
type Spec struct {
	Rules []Rule
}

// Rule is one semicolon-separated rule. A rule has at most one type.
type Rule struct {
	Selector    *jsonpath.Path
	Type        *typing.Type
	Checks      []*Check
	Conditional bool // type?: apply only when the value matches this type
}

// Parse parses an annotation, expanding named type references from defs.
func Parse(annotation string, defs map[string]string) (*Spec, error) {
	raw := strings.TrimSpace(annotation)
	expanded, err := expandTypeRefs(defs, raw, nil)
	if err != nil {
		return nil, err
	}

	parts, err := splitTopLevel(expanded, ';')
	if err != nil {
		return nil, fmt.Errorf("invalid annotation %q: %w", raw, err)
	}

	spec := &Spec{}
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			return nil, fmt.Errorf("invalid annotation %q: empty rule", raw)
		}
		rule, err := parseRule(part)
		if err != nil {
			return nil, fmt.Errorf("invalid annotation %q: %w", raw, err)
		}
		spec.Rules = append(spec.Rules, rule)
	}
	if len(spec.Rules) == 0 {
		return nil, fmt.Errorf("invalid annotation %q: no rules", raw)
	}
	return spec, nil
}

func parseRule(text string) (Rule, error) {
	var rule Rule
	rest := text

	if strings.HasPrefix(rest, "$") {
		idx := findTopLevel(rest, ':')
		if idx < 0 {
			return rule, fmt.Errorf("selector %q must be followed by ':'", rest)
		}
		selector, err := jsonpath.Parse(strings.TrimSpace(rest[:idx]))
		if err != nil {
			return rule, err
		}
		rule.Selector = selector
		rest = rest[idx+1:]
	}

	parts, err := splitTopLevelSeq(rest, "||")
	if err != nil {
		return rule, err
	}
	if len(parts) > 1 {
		return rule, fmt.Errorf("multi-type unions ('||') are not supported; use one type per rule")
	}

	typ, checks, conditional, err := parseTypeAndChecks(rest)
	if err != nil {
		return rule, err
	}
	if conditional && rule.Selector != nil {
		return rule, fmt.Errorf("conditional type '?' is only supported for whole-field rules")
	}
	rule.Type = typ
	rule.Checks = checks
	rule.Conditional = conditional
	return rule, nil
}

func parseTypeAndChecks(text string) (*typing.Type, []*Check, bool, error) {
	text = strings.TrimSpace(text)
	if text == "" {
		return nil, nil, false, fmt.Errorf("rule has no type or checks")
	}

	if idx := findTopLevel(text, ':'); idx >= 0 {
		typeText := strings.TrimSpace(text[:idx])
		checkText := strings.TrimSpace(text[idx+1:])
		if typeText == "" {
			return nil, nil, false, fmt.Errorf("missing type before ':' in %q", text)
		}
		base, conditional := stripTypeMarker(typeText)
		t, err := typing.Parse(base)
		if err != nil {
			return nil, nil, false, fmt.Errorf("invalid type %q: %w", typeText, err)
		}
		checks, err := parseChecks(checkText)
		if err != nil {
			return nil, nil, false, err
		}
		if err := compileChecks(checks, t); err != nil {
			return nil, nil, false, err
		}
		return t, checks, conditional, nil
	}

	base, conditional := stripTypeMarker(text)
	if t, err := typing.Parse(base); err == nil {
		return t, nil, conditional, nil
	}

	name := strings.ToLower(strings.TrimSpace(base))
	if _, known := checkDefs[name]; !known && !isCombinator(name) && !isDirective(name) &&
		!strings.ContainsAny(text, "(),") {
		return nil, nil, false, fmt.Errorf("unknown type or named type %q (is types.yaml present?)", text)
	}

	checks, err := parseChecks(text)
	if err != nil {
		return nil, nil, false, fmt.Errorf("invalid type or checks %q: %w", text, err)
	}
	if err := compileChecks(checks, nil); err != nil {
		return nil, nil, false, err
	}
	return nil, checks, false, nil
}

// stripTypeMarker removes a trailing '?' (conditional type rule) or '!'
// (explicitly strict) from a type expression.
func stripTypeMarker(s string) (string, bool) {
	s = strings.TrimSpace(s)
	if strings.HasSuffix(s, "?") {
		return strings.TrimSpace(strings.TrimSuffix(s, "?")), true
	}
	if strings.HasSuffix(s, "!") {
		return strings.TrimSpace(strings.TrimSuffix(s, "!")), false
	}
	return s, false
}

func expandTypeRefs(defs map[string]string, s string, seen map[string]bool) (string, error) {
	s = strings.TrimSpace(s)
	explicit := strings.HasPrefix(s, "#")

	name := s
	if explicit {
		name = strings.TrimSpace(strings.TrimPrefix(s, "#"))
	}

	def, exists := defs[name]
	if !explicit && !exists {
		return s, nil
	}
	if name == "" || !isTypeRefName(name) {
		return "", fmt.Errorf("invalid type reference %q", s)
	}
	if seen == nil {
		seen = make(map[string]bool)
	}
	if seen[name] {
		return "", fmt.Errorf("cyclic type reference %s", name)
	}
	if !exists {
		return "", fmt.Errorf("unknown type reference %q", s)
	}
	seen[name] = true
	return expandTypeRefs(defs, def, seen)
}

func isTypeRefName(s string) bool {
	for i := 0; i < len(s); i++ {
		c := s[i]
		if !(c == '_' || c == '-' || c == '.' ||
			(c >= 'a' && c <= 'z') ||
			(c >= 'A' && c <= 'Z') ||
			(c >= '0' && c <= '9')) {
			return false
		}
	}
	return true
}

// IsPlain reports whether spec is a single, unconditional, check-free type.
func IsPlain(spec *Spec) bool {
	if spec == nil || len(spec.Rules) != 1 {
		return false
	}
	rule := spec.Rules[0]
	return rule.Selector == nil && !rule.Conditional && len(rule.Checks) == 0 && rule.Type != nil
}

// ZeroForSpec returns the value emitted for an empty cell with this spec.
func ZeroForSpec(spec *Spec) any {
	for _, rule := range spec.Rules {
		if rule.Type == nil {
			continue
		}
		if rule.Type.Kind == typing.KindArray && HasPad(rule.Checks) {
			out := make([]any, 0, rule.Type.Length)
			fill := typing.ZeroValue(rule.Type.Elem)
			for i := 0; i < rule.Type.Length; i++ {
				out = append(out, fill)
			}
			return out
		}
		return typing.ZeroValue(rule.Type)
	}
	return ""
}
