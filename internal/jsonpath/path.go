package jsonpath

import (
	"fmt"
	"sort"
	"strconv"
	"strings"
)

// This file implements the subset of RFC 9535 JSONPath used by field
// annotations. It deliberately excludes filter selectors, function
// extensions and the current-node identifier "@". Evaluation returns a
// nodelist whose entries remember their location so the caller can write
// converted values back into the source document.

type jpSelectorKind int

const (
	jpName jpSelectorKind = iota
	jpIndex
	jpWildcard
	jpSlice
)

type jpSelector struct {
	kind                      jpSelectorKind
	name                      string
	index                     int
	start, end, step          int
	hasStart, hasEnd, hasStep bool
}

type jpSegment struct {
	descendant bool
	selectors  []jpSelector
}

type Path struct {
	segments []jpSegment
}

// IsRoot reports whether the path selects the document root ($).
func (p *Path) IsRoot() bool {
	return p == nil || len(p.segments) == 0
}

// Step identifies one step along a concrete document location.
type Step struct {
	Name    string
	Index   int
	IsIndex bool
}

// Node is a matched value together with its location from the root.
type Node struct {
	Value any
	Path  []Step
}

type jpParser struct {
	s   string
	pos int
}

func Parse(src string) (*Path, error) {
	s := strings.TrimSpace(src)
	if s == "" || s[0] != '$' {
		return nil, fmt.Errorf("JSONPath must start with '$': %q", src)
	}

	p := &jpParser{s: s, pos: 1}
	var segments []jpSegment

	for p.pos < len(p.s) {
		switch p.s[p.pos] {
		case '.':
			p.pos++
			descendant := false
			if p.pos < len(p.s) && p.s[p.pos] == '.' {
				descendant = true
				p.pos++
			}
			if p.pos >= len(p.s) {
				return nil, fmt.Errorf("invalid JSONPath %q: missing name after '.'", src)
			}

			if p.s[p.pos] == '[' {
				if !descendant {
					return nil, fmt.Errorf("invalid JSONPath %q: '.' cannot be followed by '['", src)
				}
				sels, err := p.parseBracket()
				if err != nil {
					return nil, err
				}
				segments = append(segments, jpSegment{descendant: true, selectors: sels})
				continue
			}

			if p.s[p.pos] == '*' {
				p.pos++
				segments = append(segments, jpSegment{
					descendant: descendant,
					selectors:  []jpSelector{{kind: jpWildcard}},
				})
				continue
			}

			start := p.pos
			for p.pos < len(p.s) && p.s[p.pos] != '.' && p.s[p.pos] != '[' {
				p.pos++
			}
			if start == p.pos {
				return nil, fmt.Errorf("invalid JSONPath %q: empty field name", src)
			}
			segments = append(segments, jpSegment{
				descendant: descendant,
				selectors:  []jpSelector{{kind: jpName, name: p.s[start:p.pos]}},
			})

		case '[':
			sels, err := p.parseBracket()
			if err != nil {
				return nil, err
			}
			segments = append(segments, jpSegment{selectors: sels})

		default:
			return nil, fmt.Errorf("invalid JSONPath %q: unexpected %q", src, p.s[p.pos])
		}
	}

	return &Path{segments: segments}, nil
}

func (p *jpParser) parseBracket() ([]jpSelector, error) {
	if p.pos >= len(p.s) || p.s[p.pos] != '[' {
		return nil, fmt.Errorf("expected '['")
	}
	p.pos++
	start := p.pos
	depth := 1
	var quote byte

	for p.pos < len(p.s) {
		c := p.s[p.pos]
		if quote != 0 {
			if c == '\\' {
				p.pos += 2
				continue
			}
			if c == quote {
				quote = 0
			}
			p.pos++
			continue
		}
		switch c {
		case '\'', '"':
			quote = c
		case '[':
			depth++
		case ']':
			depth--
			if depth == 0 {
				content := p.s[start:p.pos]
				p.pos++
				return parseBracketSelectors(content)
			}
		}
		p.pos++
	}
	return nil, fmt.Errorf("invalid JSONPath %q: unclosed '['", p.s)
}

func splitSelectors(s string) ([]string, error) {
	var parts []string
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
		case ',':
			parts = append(parts, s[start:i])
			start = i + 1
		}
	}
	if quote != 0 {
		return nil, fmt.Errorf("unterminated quote")
	}
	parts = append(parts, s[start:])
	return parts, nil
}

func parseBracketSelectors(content string) ([]jpSelector, error) {
	parts, err := splitSelectors(content)
	if err != nil {
		return nil, err
	}
	if len(parts) == 0 {
		return nil, fmt.Errorf("empty bracket selector")
	}

	var sels []jpSelector
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			return nil, fmt.Errorf("empty bracket selector in %q", content)
		}

		if part == "*" {
			sels = append(sels, jpSelector{kind: jpWildcard})
			continue
		}
		if part[0] == '\'' || part[0] == '"' {
			name, err := unquoteJPName(part)
			if err != nil {
				return nil, err
			}
			sels = append(sels, jpSelector{kind: jpName, name: name})
			continue
		}
		if strings.Contains(part, ":") {
			sel, err := parseSliceSelector(part)
			if err != nil {
				return nil, err
			}
			sels = append(sels, sel)
			continue
		}
		n, err := strconv.Atoi(part)
		if err != nil {
			return nil, fmt.Errorf("invalid array index %q", part)
		}
		sels = append(sels, jpSelector{kind: jpIndex, index: n})
	}
	return sels, nil
}

func parseSliceSelector(part string) (jpSelector, error) {
	fields := strings.Split(part, ":")
	if len(fields) < 2 || len(fields) > 3 {
		return jpSelector{}, fmt.Errorf("invalid array slice %q", part)
	}

	sel := jpSelector{kind: jpSlice}
	var err error
	sel.start, sel.hasStart, err = parseIntField(fields[0])
	if err != nil {
		return jpSelector{}, fmt.Errorf("invalid slice start in %q: %w", part, err)
	}
	sel.end, sel.hasEnd, err = parseIntField(fields[1])
	if err != nil {
		return jpSelector{}, fmt.Errorf("invalid slice end in %q: %w", part, err)
	}
	if len(fields) == 3 {
		sel.step, sel.hasStep, err = parseIntField(fields[2])
		if err != nil {
			return jpSelector{}, fmt.Errorf("invalid slice step in %q: %w", part, err)
		}
		if sel.hasStep && sel.step == 0 {
			return jpSelector{}, fmt.Errorf("array slice step cannot be zero in %q", part)
		}
	}
	return sel, nil
}

func parseIntField(s string) (int, bool, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, false, nil
	}
	n, err := strconv.Atoi(s)
	if err != nil {
		return 0, false, err
	}
	return n, true, nil
}

func unquoteJPName(s string) (string, error) {
	if len(s) < 2 || s[len(s)-1] != s[0] {
		return "", fmt.Errorf("invalid quoted name %q", s)
	}
	if s[0] == '"' {
		v, err := strconv.Unquote(s)
		if err != nil {
			return "", fmt.Errorf("invalid quoted name %q: %w", s, err)
		}
		return v, nil
	}

	var b strings.Builder
	for i := 1; i < len(s)-1; i++ {
		c := s[i]
		if c != '\\' {
			b.WriteByte(c)
			continue
		}
		i++
		if i >= len(s)-1 {
			return "", fmt.Errorf("invalid escape in quoted name %q", s)
		}
		switch s[i] {
		case '\\', '\'', '/', '"':
			b.WriteByte(s[i])
		case 'b':
			b.WriteByte('\b')
		case 'f':
			b.WriteByte('\f')
		case 'n':
			b.WriteByte('\n')
		case 'r':
			b.WriteByte('\r')
		case 't':
			b.WriteByte('\t')
		case 'u':
			if i+4 >= len(s) {
				return "", fmt.Errorf("invalid unicode escape in quoted name %q", s)
			}
			r, err := strconv.ParseUint(s[i+1:i+5], 16, 16)
			if err != nil {
				return "", fmt.Errorf("invalid unicode escape in quoted name %q: %w", s, err)
			}
			b.WriteRune(rune(r))
			i += 4
		default:
			return "", fmt.Errorf("invalid escape %q in quoted name", s[i])
		}
	}
	return b.String(), nil
}

func (p *Path) Eval(root any) []Node {
	nodes := []Node{{Value: root}}
	for _, seg := range p.segments {
		next := make([]Node, 0, len(nodes))
		for _, node := range nodes {
			if seg.descendant {
				next = appendDescendants(next, node.Value, node.Path, seg.selectors)
			} else {
				next = appendSelectors(next, node.Value, node.Path, seg.selectors)
			}
		}
		nodes = next
	}
	return nodes
}

func appendSelectors(out []Node, v any, path []Step, sels []jpSelector) []Node {
	for _, sel := range sels {
		switch sel.kind {
		case jpName:
			m, ok := v.(map[string]any)
			if !ok {
				continue
			}
			child, ok := m[sel.name]
			if !ok {
				continue
			}
			out = append(out, Node{Value: child, Path: appendStep(path, Step{Name: sel.name})})

		case jpIndex:
			arr, ok := v.([]any)
			if !ok {
				continue
			}
			idx := sel.index
			if idx < 0 {
				idx += len(arr)
			}
			if idx < 0 || idx >= len(arr) {
				continue
			}
			out = append(out, Node{Value: arr[idx], Path: appendStep(path, Step{Index: idx, IsIndex: true})})

		case jpWildcard:
			switch x := v.(type) {
			case []any:
				for i, child := range x {
					out = append(out, Node{Value: child, Path: appendStep(path, Step{Index: i, IsIndex: true})})
				}
			case map[string]any:
				for _, key := range sortedKeys(x) {
					out = append(out, Node{Value: x[key], Path: appendStep(path, Step{Name: key})})
				}
			}

		case jpSlice:
			arr, ok := v.([]any)
			if !ok {
				continue
			}
			for _, idx := range sliceIndices(len(arr), sel) {
				out = append(out, Node{Value: arr[idx], Path: appendStep(path, Step{Index: idx, IsIndex: true})})
			}
		}
	}
	return out
}

func appendDescendants(out []Node, v any, path []Step, sels []jpSelector) []Node {
	out = appendSelectors(out, v, path, sels)
	switch x := v.(type) {
	case []any:
		for i, child := range x {
			out = appendDescendants(out, child, appendStep(path, Step{Index: i, IsIndex: true}), sels)
		}
	case map[string]any:
		for _, key := range sortedKeys(x) {
			out = appendDescendants(out, x[key], appendStep(path, Step{Name: key}), sels)
		}
	}
	return out
}

func appendStep(path []Step, step Step) []Step {
	out := make([]Step, len(path)+1)
	copy(out, path)
	out[len(path)] = step
	return out
}

func sortedKeys(m map[string]any) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	return keys
}

// sliceIndices implements Python/RFC 9535 slice semantics.
func sliceIndices(n int, sel jpSelector) []int {
	step := 1
	if sel.hasStep {
		step = sel.step
	}
	if step == 0 {
		return nil
	}

	start := 0
	if step < 0 {
		start = n - 1
	}
	if sel.hasStart {
		start = sel.start
		if start < 0 {
			start += n
			if start < 0 {
				if step < 0 {
					start = -1
				} else {
					start = 0
				}
			}
		} else if start >= n {
			if step < 0 {
				start = n - 1
			} else {
				start = n
			}
		}
	}

	end := n
	if step < 0 {
		end = -1
	}
	if sel.hasEnd {
		end = sel.end
		if end < 0 {
			end += n
			if end < 0 {
				if step < 0 {
					end = -1
				} else {
					end = 0
				}
			}
		} else if end >= n {
			if step < 0 {
				end = n - 1
			} else {
				end = n
			}
		}
	}

	var out []int
	if step > 0 {
		for i := start; i < end; i += step {
			out = append(out, i)
		}
	} else {
		for i := start; i > end; i += step {
			out = append(out, i)
		}
	}
	return out
}

// setAtPath writes value at path inside root and returns the (possibly new) root.
func SetAtPath(root any, path []Step, value any) (any, error) {
	if len(path) == 0 {
		return value, nil
	}

	cur := root
	for i, step := range path {
		last := i == len(path)-1
		if step.IsIndex {
			arr, ok := cur.([]any)
			if !ok {
				return root, fmt.Errorf("cannot index non-array at step %d", i)
			}
			if step.Index < 0 || step.Index >= len(arr) {
				return root, fmt.Errorf("index %d out of range at step %d", step.Index, i)
			}
			if last {
				arr[step.Index] = value
				return root, nil
			}
			cur = arr[step.Index]
			continue
		}

		m, ok := cur.(map[string]any)
		if !ok {
			return root, fmt.Errorf("cannot select field %q from non-object at step %d", step.Name, i)
		}
		if last {
			m[step.Name] = value
			return root, nil
		}
		next, ok := m[step.Name]
		if !ok {
			return root, fmt.Errorf("field %q missing at step %d", step.Name, i)
		}
		cur = next
	}
	return root, nil
}
