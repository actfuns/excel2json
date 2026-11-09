// Package parseutil provides delimited-expression splitting helpers that are
// aware of nested brackets, parentheses and quotes.
package schema

import (
	"fmt"
	"strings"
)

// splitTopLevel splits s on sep, ignoring separators inside (), [], {} or
// quoted strings.
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

// splitTopLevelSeq splits s on a multi-character separator at the top level.
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

// findTopLevel returns the index of the first target byte at nesting depth 0,
// or -1.
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
