package convert

import (
	"fmt"
	"strings"

	"github.com/actfuns/excel2json/internal/core"
	"github.com/actfuns/excel2json/internal/schema"
	"github.com/actfuns/excel2json/internal/typing"
)

type rowContext struct {
	sheet *core.Sheet
	row   []string
}

func (e *Engine) matchGuards(conditions []schema.Condition, cell string, row *rowContext) (bool, error) {
	for _, cond := range conditions {
		ok, err := e.matchCondition(cond, cell, row)
		if err != nil {
			return false, err
		}
		if !ok {
			return false, nil
		}
	}
	return true, nil
}

func (e *Engine) matchCondition(cond schema.Condition, cell string, row *rowContext) (bool, error) {
	switch cond.Kind {
	case schema.CondAnd:
		for _, sub := range cond.Subs {
			ok, err := e.matchCondition(sub, cell, row)
			if err != nil || !ok {
				return false, err
			}
		}
		return true, nil
	case schema.CondOr:
		var errs []string
		for _, sub := range cond.Subs {
			ok, err := e.matchCondition(sub, cell, row)
			if err != nil {
				errs = append(errs, err.Error())
				continue
			}
			if ok {
				return true, nil
			}
		}
		if len(errs) > 0 {
			return false, fmt.Errorf("or: %s", strings.Join(errs, "; "))
		}
		return false, nil
	case schema.CondNot:
		for _, sub := range cond.Subs {
			ok, err := e.matchCondition(sub, cell, row)
			if err != nil {
				return false, err
			}
			if !ok {
				return true, nil
			}
		}
		return false, nil
	}

	values, err := e.resolveGuardRef(cond.Ref, cell, row)
	if err != nil {
		return false, err
	}
	if len(values) == 0 {
		return false, nil
	}
	for _, value := range values {
		if cond.Kind == schema.CondType {
			ok := typing.MatchJSON(value, cond.TypeExpr)
			if cond.Op == schema.OpNe {
				ok = !ok
			}
			if ok {
				return true, nil
			}
			continue
		}
		ok, err := schema.CompareValue(value, cond.Op, cond.Literal)
		if err != nil {
			return false, err
		}
		if ok {
			return true, nil
		}
	}
	return false, nil
}

func (e *Engine) resolveGuardRef(ref schema.Ref, cell string, row *rowContext) ([]any, error) {
	if row == nil {
		return nil, fmt.Errorf("conditions require row context")
	}

	raw := cell
	if ref.Column != "" {
		idx := row.sheet.ColumnIndex(ref.Column)
		if idx < 0 {
			return nil, fmt.Errorf("condition column %q not found", ref.Column)
		}
		raw = ""
		if idx < len(row.row) {
			raw = strings.TrimSpace(row.row[idx])
		}
	} else {
		raw = strings.TrimSpace(raw)
	}

	if ref.Path == nil || ref.Path.IsRoot() {
		return []any{decodeGuardRaw(raw)}, nil
	}

	if raw == "" {
		return nil, nil
	}
	decoded, err := typing.DecodeJSONCell(raw)
	if err != nil {
		return nil, fmt.Errorf("condition reference %q is not valid JSON", raw)
	}
	nodes := ref.Path.Eval(decoded)
	out := make([]any, len(nodes))
	for i, node := range nodes {
		out[i] = node.Value
	}
	return out, nil
}

func decodeGuardRaw(raw string) any {
	if decoded, err := typing.DecodeJSONCell(raw); err == nil {
		return decoded
	}
	switch strings.ToLower(raw) {
	case "true":
		return true
	case "false":
		return false
	default:
		return raw
	}
}
