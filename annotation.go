package main

import (
	"fmt"
	"strings"
)

// A fieldSpec is the parsed form of a type-row annotation.
//
//	annotation = rule { ";" rule }
//	rule       = [ selector ":" ] [ type [ "!" | "?" ] [ ":" checks ] ]
//	checks     = check { "," check }
//
// A rule has at most one type. Multiple rules are separated by ";" and are
// applied to the same field.
type fieldSpec struct {
	rules []fieldRule
}

type fieldRule struct {
	selector    *jsonPath
	typ         *goType
	checks      []*checkExpr
	conditional bool // type?: apply only when the value matches this type
}

type cachedFieldSpec struct {
	spec *fieldSpec
	err  error
}

func (e *Exporter) parseFieldSpecCached(annotation string) (*fieldSpec, error) {
	if v, ok := e.specCache.Load(annotation); ok {
		c := v.(cachedFieldSpec)
		return c.spec, c.err
	}
	spec, err := e.parseFieldSpec(annotation)
	e.specCache.Store(annotation, cachedFieldSpec{spec: spec, err: err})
	return spec, err
}

func (e *Exporter) parseFieldSpec(annotation string) (*fieldSpec, error) {
	raw := strings.TrimSpace(annotation)
	expanded, err := e.expandTypeRefs(raw, nil)
	if err != nil {
		return nil, err
	}

	parts, err := splitTopLevel(expanded, ';')
	if err != nil {
		return nil, fmt.Errorf("invalid annotation %q: %w", raw, err)
	}

	spec := &fieldSpec{}
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			return nil, fmt.Errorf("invalid annotation %q: empty rule", raw)
		}
		rule, err := e.parseRule(part)
		if err != nil {
			return nil, fmt.Errorf("invalid annotation %q: %w", raw, err)
		}
		spec.rules = append(spec.rules, rule)
	}
	if len(spec.rules) == 0 {
		return nil, fmt.Errorf("invalid annotation %q: no rules", raw)
	}
	return spec, nil
}

func (e *Exporter) parseRule(text string) (fieldRule, error) {
	var rule fieldRule
	rest := text

	if strings.HasPrefix(rest, "$") {
		idx := findTopLevel(rest, ':')
		if idx < 0 {
			return rule, fmt.Errorf("selector %q must be followed by ':'", rest)
		}
		selector, err := parseJSONPath(strings.TrimSpace(rest[:idx]))
		if err != nil {
			return rule, err
		}
		rule.selector = selector
		rest = rest[idx+1:]
	}

	parts, err := splitTopLevelSeq(rest, "||")
	if err != nil {
		return rule, err
	}
	if len(parts) > 1 {
		return rule, fmt.Errorf("multi-type unions ('||') are not supported; use one type per rule")
	}

	typ, checks, conditional, err := e.parseTypeAndChecks(rest)
	if err != nil {
		return rule, err
	}
	if conditional && rule.selector != nil {
		return rule, fmt.Errorf("conditional type '?' is only supported for whole-field rules")
	}
	rule.typ = typ
	rule.checks = checks
	rule.conditional = conditional
	return rule, nil
}

func (e *Exporter) parseTypeAndChecks(text string) (*goType, []*checkExpr, bool, error) {
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
		t, err := parseType(base)
		if err != nil {
			return nil, nil, false, fmt.Errorf("invalid type %q: %w", typeText, err)
		}
		checks, err := e.parseChecks(checkText)
		if err != nil {
			return nil, nil, false, err
		}
		if err := e.compileChecks(checks, t); err != nil {
			return nil, nil, false, err
		}
		return t, checks, conditional, nil
	}

	base, conditional := stripTypeMarker(text)
	if t, err := parseType(base); err == nil {
		return t, nil, conditional, nil
	}

	name := strings.ToLower(strings.TrimSpace(base))
	if _, known := checkDefs[name]; !known && !isCombinator(name) && !isDirective(name) &&
		!strings.ContainsAny(text, "(),") {
		return nil, nil, false, fmt.Errorf("unknown type or named type %q (is types.yaml present?)", text)
	}

	checks, err := e.parseChecks(text)
	if err != nil {
		return nil, nil, false, fmt.Errorf("invalid type or checks %q: %w", text, err)
	}
	if err := e.compileChecks(checks, nil); err != nil {
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

func (e *Exporter) parseChecks(text string) ([]*checkExpr, error) {
	text = strings.TrimSpace(text)
	if text == "" {
		return nil, nil
	}
	parts, err := splitTopLevel(text, ',')
	if err != nil {
		return nil, fmt.Errorf("invalid checks %q: %w", text, err)
	}
	var checks []*checkExpr
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			return nil, fmt.Errorf("invalid checks %q: empty check", text)
		}
		ch, err := parseCheck(part)
		if err != nil {
			return nil, err
		}
		checks = append(checks, ch)
	}
	return checks, nil
}

func parseCheck(text string) (*checkExpr, error) {
	text = strings.TrimSpace(text)
	i := 0
	for i < len(text) && isCheckNameChar(text[i]) {
		i++
	}
	if i == 0 {
		return nil, fmt.Errorf("invalid check %q", text)
	}

	ch := &checkExpr{name: strings.ToLower(text[:i])}
	rest := text[i:]
	if rest == "" {
		return ch, nil
	}
	if rest[0] != '(' {
		return nil, fmt.Errorf("invalid check %q: expected '(' after %s", text, ch.name)
	}
	if !strings.HasSuffix(rest, ")") {
		return nil, fmt.Errorf("invalid check %q: missing ')'", text)
	}

	inner := rest[1 : len(rest)-1]
	if strings.TrimSpace(inner) != "" {
		args, err := splitTopLevel(inner, ',')
		if err != nil {
			return nil, fmt.Errorf("invalid check %q: %w", text, err)
		}
		for _, arg := range args {
			arg = strings.TrimSpace(arg)
			if arg == "" {
				return nil, fmt.Errorf("invalid check %q: empty argument", text)
			}
			ch.args = append(ch.args, arg)
		}
	}

	if isCombinator(ch.name) {
		for _, arg := range ch.args {
			sub, err := parseCheck(arg)
			if err != nil {
				return nil, fmt.Errorf("invalid %s operand %q: %w", ch.name, arg, err)
			}
			ch.subs = append(ch.subs, sub)
		}
		ch.args = nil
	}
	return ch, nil
}

func isCheckNameChar(c byte) bool {
	return c == '_' ||
		(c >= 'a' && c <= 'z') ||
		(c >= 'A' && c <= 'Z') ||
		(c >= '0' && c <= '9')
}

func (e *Exporter) compileChecks(checks []*checkExpr, t *goType) error {
	padSeen := false
	for _, ch := range checks {
		if isDirective(ch.name) {
			if padSeen {
				return fmt.Errorf("duplicate pad directive")
			}
			padSeen = true
			def, ok := checkDefs[ch.name]
			if !ok {
				return fmt.Errorf("unknown check %q", ch.name)
			}
			if len(ch.args) > def.maxArgs {
				return fmt.Errorf("pad expects at most %d argument(s), got %d", def.maxArgs, len(ch.args))
			}
			if t == nil || t.kind != kindArray {
				return fmt.Errorf("pad requires a fixed-size array type")
			}
			continue
		}
		if err := e.compileCheck(ch, t); err != nil {
			return err
		}
	}
	return nil
}

func (e *Exporter) expandTypeRefs(s string, seen map[string]bool) (string, error) {
	s = strings.TrimSpace(s)
	explicit := strings.HasPrefix(s, "#")

	name := s
	if explicit {
		name = strings.TrimSpace(strings.TrimPrefix(s, "#"))
	}

	def, exists := e.opts.TypeDefs[name]
	if !explicit && !exists {
		// Not a named type: leave built-in types and plain checks untouched.
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
	return e.expandTypeRefs(def, seen)
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

func splitTopLevel(s string, sep byte) ([]string, error) {
	var parts []string
	depth := 0
	var quote byte
	start := 0

	for i := 0; i < len(s); i++ {
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
		case '(', '[', '{':
			depth++
		case ')', ']', '}':
			if depth == 0 {
				return nil, fmt.Errorf("unbalanced %q", c)
			}
			depth--
		default:
			if c == sep && depth == 0 {
				parts = append(parts, s[start:i])
				start = i + 1
			}
		}
	}
	if quote != 0 {
		return nil, fmt.Errorf("unterminated quote")
	}
	if depth != 0 {
		return nil, fmt.Errorf("unbalanced brackets")
	}
	parts = append(parts, s[start:])
	return parts, nil
}

func splitTopLevelSeq(s, sep string) ([]string, error) {
	if sep == "" {
		return nil, fmt.Errorf("empty separator")
	}
	var parts []string
	depth := 0
	var quote byte
	start := 0

	for i := 0; i < len(s); i++ {
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
		if depth == 0 && strings.HasPrefix(s[i:], sep) {
			parts = append(parts, s[start:i])
			i += len(sep) - 1
			start = i + 1
			continue
		}
		switch c {
		case '\'', '"':
			quote = c
		case '(', '[', '{':
			depth++
		case ')', ']', '}':
			if depth == 0 {
				return nil, fmt.Errorf("unbalanced %q", c)
			}
			depth--
		}
	}
	if quote != 0 {
		return nil, fmt.Errorf("unterminated quote")
	}
	if depth != 0 {
		return nil, fmt.Errorf("unbalanced brackets")
	}
	parts = append(parts, s[start:])
	return parts, nil
}

func findTopLevel(s string, target byte) int {
	depth := 0
	var quote byte

	for i := 0; i < len(s); i++ {
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
		case '(', '[', '{':
			depth++
		case ')', ']', '}':
			if depth > 0 {
				depth--
			}
		default:
			if c == target && depth == 0 {
				return i
			}
		}
	}
	return -1
}
