package main

import (
	"encoding/json"
	"fmt"
	"math/big"
	"net/mail"
	"net/url"
	"regexp"
	"strconv"
	"strings"
	"sync"
	"time"
	"unicode"
)

// checkExpr is a parsed validator expression. Combinators (and/or/not) keep
// their operands in subs; ordinary validators keep raw literal arguments.
type checkExpr struct {
	name string
	args []string
	subs []*checkExpr
}

func isCombinator(name string) bool {
	return name == "and" || name == "or" || name == "not"
}

func isDirective(name string) bool {
	return name == "pad"
}

type checkKind int

const (
	ckAny checkKind = iota
	ckNumber
	ckString
	ckCollection
	ckScalar
	ckDate
	ckArray
)

type checkDef struct {
	minArgs int
	maxArgs int // -1 = unlimited
	kind    checkKind
}

var checkDefs = map[string]checkDef{
	"gt":         {1, 1, ckNumber},
	"gte":        {1, 1, ckNumber},
	"lt":         {1, 1, ckNumber},
	"lte":        {1, 1, ckNumber},
	"between":    {2, 2, ckNumber},
	"eq":         {1, 1, ckScalar},
	"ne":         {1, 1, ckScalar},
	"positive":   {0, 0, ckNumber},
	"negative":   {0, 0, ckNumber},
	"nonzero":    {0, 0, ckNumber},
	"multipleof": {1, 1, ckNumber},

	"len":      {1, 2, ckCollection},
	"minlen":   {1, 1, ckCollection},
	"maxlen":   {1, 1, ckCollection},
	"email":    {0, 0, ckString},
	"phone":    {0, 1, ckString},
	"url":      {0, 0, ckString},
	"uuid":     {0, 0, ckString},
	"regex":    {1, 1, ckString},
	"prefix":   {1, 1, ckString},
	"suffix":   {1, 1, ckString},
	"contains": {1, 1, ckString},
	"alpha":    {0, 0, ckString},
	"alnum":    {0, 0, ckString},
	"numeric":  {0, 0, ckString},

	"oneof":    {1, -1, ckScalar},
	"in":       {1, -1, ckScalar},
	"notoneof": {1, -1, ckScalar},
	"notin":    {1, -1, ckScalar},

	"unique":   {0, 0, ckArray},
	"nonempty": {0, 0, ckCollection},
	"ref":      {1, 1, ckScalar},
	"after":    {1, 1, ckDate},
	"before":   {1, 1, ckDate},
	"pad":      {0, 1, ckArray},
}

func checkKindMatches(kind checkKind, t *goType) bool {
	if t == nil || t.kind == kindAny || kind == ckAny {
		return true
	}
	switch kind {
	case ckCollection:
		return t.kind == kindString || t.kind == kindSlice || t.kind == kindArray || t.kind == kindMap
	case ckArray:
		return t.kind == kindSlice || t.kind == kindArray
	}
	// Scalar-style checks on arrays are applied element-wise.
	if (t.kind == kindSlice || t.kind == kindArray) && t.elem != nil {
		return checkKindMatches(kind, t.elem)
	}
	switch kind {
	case ckNumber:
		return t.kind == kindInt || t.kind == kindUint || t.kind == kindFloat
	case ckString:
		return t.kind == kindString
	case ckScalar:
		return t.kind == kindBool || t.kind == kindString || t.kind == kindInt || t.kind == kindUint || t.kind == kindFloat || t.kind == kindDate
	case ckDate:
		return t.kind == kindDate
	}
	return false
}

func (e *Exporter) compileCheck(ch *checkExpr, t *goType) error {
	if isDirective(ch.name) {
		return fmt.Errorf("%s may only appear directly in a rule, not inside combinators", ch.name)
	}
	if isCombinator(ch.name) {
		if len(ch.subs) == 0 {
			return fmt.Errorf("%s requires at least one check", ch.name)
		}
		for _, sub := range ch.subs {
			if err := e.compileCheck(sub, t); err != nil {
				return err
			}
		}
		return nil
	}

	def, ok := checkDefs[ch.name]
	if !ok {
		return fmt.Errorf("unknown check %q", ch.name)
	}
	if len(ch.args) < def.minArgs || (def.maxArgs >= 0 && len(ch.args) > def.maxArgs) {
		if def.maxArgs < 0 {
			return fmt.Errorf("%s expects at least %d argument(s), got %d", ch.name, def.minArgs, len(ch.args))
		}
		if def.minArgs == def.maxArgs {
			return fmt.Errorf("%s expects %d argument(s), got %d", ch.name, def.minArgs, len(ch.args))
		}
		return fmt.Errorf("%s expects %d..%d argument(s), got %d", ch.name, def.minArgs, def.maxArgs, len(ch.args))
	}
	if !checkKindMatches(def.kind, t) {
		return fmt.Errorf("check %q does not apply to type %s", ch.name, t)
	}
	return nil
}

func (e *Exporter) runChecks(v any, checks []*checkExpr) error {
	arr, isArray := v.([]any)
	for _, ch := range checks {
		if isDirective(ch.name) {
			continue
		}
		if isArray && !isCollectionCheckExpr(ch) {
			for i, item := range arr {
				if err := e.runChecks(item, []*checkExpr{ch}); err != nil {
					return fmt.Errorf("index %d: %w", i, err)
				}
			}
			continue
		}
		if err := e.runCheck(v, ch); err != nil {
			return err
		}
	}
	return nil
}

// isCollectionCheckExpr reports whether a check (or any operand of a
// combinator) targets the collection as a whole rather than its elements.
func isCollectionCheckExpr(ch *checkExpr) bool {
	if isCombinator(ch.name) {
		for _, sub := range ch.subs {
			if isCollectionCheckExpr(sub) {
				return true
			}
		}
		return false
	}
	switch ch.name {
	case "len", "minlen", "maxlen", "unique", "nonempty":
		return true
	}
	return false
}

func (e *Exporter) runCheck(v any, ch *checkExpr) error {
	switch ch.name {
	case "and":
		for _, sub := range ch.subs {
			if err := e.runCheck(v, sub); err != nil {
				return err
			}
		}
		return nil
	case "or":
		var errs []string
		for _, sub := range ch.subs {
			if err := e.runCheck(v, sub); err == nil {
				return nil
			} else {
				errs = append(errs, err.Error())
			}
		}
		return fmt.Errorf("or: none matched (%s)", strings.Join(errs, "; "))
	case "not":
		allPass := true
		for _, sub := range ch.subs {
			if err := e.runCheck(v, sub); err != nil {
				allPass = false
			}
		}
		if allPass {
			return fmt.Errorf("not: condition matched")
		}
		return nil
	}

	def := checkDefs[ch.name]
	if def.kind == ckNumber || def.kind == ckScalar {
		if _, ok := toBigFloat(v); !ok {
			if def.kind == ckNumber {
				return fmt.Errorf("%s expects a number, got %T", ch.name, v)
			}
		}
	}

	switch ch.name {
	case "gt", "gte", "lt", "lte":
		cmp, err := compareNumericArg(v, ch.args[0])
		if err != nil {
			return fmt.Errorf("%s: %w", ch.name, err)
		}
		ok := cmp > 0
		if ch.name == "gte" {
			ok = cmp >= 0
		} else if ch.name == "lt" {
			ok = cmp < 0
		} else if ch.name == "lte" {
			ok = cmp <= 0
		}
		if !ok {
			return fmt.Errorf("%s(%s) failed for %s", ch.name, ch.args[0], canonicalString(v))
		}
		return nil

	case "between":
		lo, err := compareNumericArg(v, ch.args[0])
		if err != nil {
			return fmt.Errorf("between: %w", err)
		}
		hi, err := compareNumericArg(v, ch.args[1])
		if err != nil {
			return fmt.Errorf("between: %w", err)
		}
		if lo < 0 || hi > 0 {
			return fmt.Errorf("between(%s,%s) failed for %s", ch.args[0], ch.args[1], canonicalString(v))
		}
		return nil

	case "eq", "ne":
		equal, err := equalLiteral(v, ch.args[0])
		if err != nil {
			return fmt.Errorf("%s: %w", ch.name, err)
		}
		if (ch.name == "eq") != equal {
			return fmt.Errorf("%s(%s) failed for %s", ch.name, ch.args[0], canonicalString(v))
		}
		return nil

	case "positive", "negative", "nonzero":
		cmp, err := compareNumericArg(v, "0")
		if err != nil {
			return fmt.Errorf("%s: %w", ch.name, err)
		}
		ok := cmp != 0
		if ch.name == "positive" {
			ok = cmp > 0
		} else if ch.name == "negative" {
			ok = cmp < 0
		}
		if !ok {
			return fmt.Errorf("%s failed for %s", ch.name, canonicalString(v))
		}
		return nil

	case "multipleof":
		vf, ok := toBigFloat(v)
		if !ok {
			return fmt.Errorf("multipleof expects a number, got %T", v)
		}
		af, err := argBigFloat(ch.args[0])
		if err != nil {
			return fmt.Errorf("multipleof: %w", err)
		}
		if af.Sign() == 0 {
			return fmt.Errorf("multipleof divisor cannot be zero")
		}
		q := new(big.Float).Quo(vf, af)
		if !q.IsInt() {
			return fmt.Errorf("multipleof(%s) failed for %s", ch.args[0], canonicalString(v))
		}
		return nil

	case "len", "minlen", "maxlen":
		n, ok := valueLen(v)
		if !ok {
			return fmt.Errorf("%s expects string/array/map, got %T", ch.name, v)
		}
		switch ch.name {
		case "len":
			if len(ch.args) == 1 {
				want, err := strconv.Atoi(ch.args[0])
				if err != nil {
					return fmt.Errorf("len: invalid length %q", ch.args[0])
				}
				if n != want {
					return fmt.Errorf("len(%d) failed, got %d", want, n)
				}
				return nil
			}
			lo, err := strconv.Atoi(ch.args[0])
			if err != nil {
				return fmt.Errorf("len: invalid length %q", ch.args[0])
			}
			hi, err := strconv.Atoi(ch.args[1])
			if err != nil {
				return fmt.Errorf("len: invalid length %q", ch.args[1])
			}
			if n < lo || n > hi {
				return fmt.Errorf("len(%d,%d) failed, got %d", lo, hi, n)
			}
			return nil
		case "minlen":
			lo, err := strconv.Atoi(ch.args[0])
			if err != nil {
				return fmt.Errorf("minlen: invalid length %q", ch.args[0])
			}
			if n < lo {
				return fmt.Errorf("minlen(%d) failed, got %d", lo, n)
			}
			return nil
		default:
			hi, err := strconv.Atoi(ch.args[0])
			if err != nil {
				return fmt.Errorf("maxlen: invalid length %q", ch.args[0])
			}
			if n > hi {
				return fmt.Errorf("maxlen(%d) failed, got %d", hi, n)
			}
			return nil
		}

	case "email":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("email expects string, got %T", v)
		}
		addr, err := mail.ParseAddress(s)
		if err != nil || addr.Name != "" || addr.Address != s || !strings.Contains(s, "@") {
			return fmt.Errorf("email failed for %q", s)
		}
		return nil

	case "phone":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("phone expects string, got %T", v)
		}
		pattern := `^1[3-9]\d{9}$`
		if len(ch.args) == 1 {
			switch argLiteral(ch.args[0]) {
			case "", "cn":
				// keep default
			case "e164":
				pattern = `^\+?[1-9]\d{6,14}$`
			default:
				pattern = argLiteral(ch.args[0])
			}
		}
		re, err := compileRegex(pattern)
		if err != nil {
			return fmt.Errorf("phone: %w", err)
		}
		if !re.MatchString(s) {
			return fmt.Errorf("phone failed for %q", s)
		}
		return nil

	case "url":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("url expects string, got %T", v)
		}
		u, err := url.ParseRequestURI(s)
		if err != nil || u.Scheme == "" || u.Host == "" {
			return fmt.Errorf("url failed for %q", s)
		}
		return nil

	case "uuid":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("uuid expects string, got %T", v)
		}
		re := regexp.MustCompile(`^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$`)
		if !re.MatchString(s) {
			return fmt.Errorf("uuid failed for %q", s)
		}
		return nil

	case "regex":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("regex expects string, got %T", v)
		}
		re, err := compileRegex(argLiteral(ch.args[0]))
		if err != nil {
			return fmt.Errorf("regex: %w", err)
		}
		if !re.MatchString(s) {
			return fmt.Errorf("regex(%s) failed for %q", ch.args[0], s)
		}
		return nil

	case "prefix", "suffix", "contains":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("%s expects string, got %T", ch.name, v)
		}
		arg := argLiteral(ch.args[0])
		var okVal bool
		switch ch.name {
		case "prefix":
			okVal = strings.HasPrefix(s, arg)
		case "suffix":
			okVal = strings.HasSuffix(s, arg)
		default:
			okVal = strings.Contains(s, arg)
		}
		if !okVal {
			return fmt.Errorf("%s(%q) failed for %q", ch.name, arg, s)
		}
		return nil

	case "alpha", "alnum", "numeric":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("%s expects string, got %T", ch.name, v)
		}
		for _, r := range s {
			switch ch.name {
			case "alpha":
				ok = unicode.IsLetter(r)
			case "alnum":
				ok = unicode.IsLetter(r) || unicode.IsDigit(r)
			default:
				ok = unicode.IsDigit(r)
			}
			if !ok {
				return fmt.Errorf("%s failed for %q", ch.name, s)
			}
		}
		return nil

	case "oneof", "in", "notoneof", "notin":
		any := false
		for _, arg := range ch.args {
			equal, err := equalLiteral(v, arg)
			if err != nil {
				return fmt.Errorf("%s: %w", ch.name, err)
			}
			if equal {
				any = true
				break
			}
		}
		if ch.name == "oneof" || ch.name == "in" {
			if !any {
				return fmt.Errorf("%s(%s) failed for %s", ch.name, strings.Join(ch.args, ","), canonicalString(v))
			}
			return nil
		}
		if any {
			return fmt.Errorf("%s(%s) failed for %s", ch.name, strings.Join(ch.args, ","), canonicalString(v))
		}
		return nil

	case "unique":
		arr, ok := v.([]any)
		if !ok {
			return fmt.Errorf("unique expects array, got %T", v)
		}
		seen := make(map[string]struct{}, len(arr))
		for _, item := range arr {
			k := canonicalString(item)
			if _, dup := seen[k]; dup {
				return fmt.Errorf("unique failed: duplicate %q", k)
			}
			seen[k] = struct{}{}
		}
		return nil

	case "nonempty":
		n, ok := valueLen(v)
		if !ok {
			return fmt.Errorf("nonempty expects string/array/map, got %T", v)
		}
		if n == 0 {
			return fmt.Errorf("nonempty failed")
		}
		return nil

	case "ref":
		return e.checkRef(v, ch.args[0])

	case "after", "before":
		dv, ok := v.(dateValue)
		if !ok {
			return fmt.Errorf("%s expects date, got %T", ch.name, v)
		}
		want, err := parseDateLiteral(ch.args[0])
		if err != nil {
			return fmt.Errorf("%s: %w", ch.name, err)
		}
		if ch.name == "after" {
			if !dv.t.After(want) {
				return fmt.Errorf("after(%s) failed for %s", ch.args[0], dv.String())
			}
		} else if !dv.t.Before(want) {
			return fmt.Errorf("before(%s) failed for %s", ch.args[0], dv.String())
		}
		return nil
	}

	return fmt.Errorf("unknown check %q", ch.name)
}

func argLiteral(s string) string {
	s = strings.TrimSpace(s)
	if v, ok := unquoteArg(s); ok {
		return v
	}
	return s
}

func unquoteArg(s string) (string, bool) {
	if len(s) >= 2 && s[0] == '"' && s[len(s)-1] == '"' {
		v, err := strconv.Unquote(s)
		if err == nil {
			return v, true
		}
	}
	if len(s) >= 2 && s[0] == '\'' && s[len(s)-1] == '\'' {
		v, err := unquoteJPName(s)
		if err == nil {
			return v, true
		}
	}
	return "", false
}

func toBigFloat(v any) (*big.Float, bool) {
	switch x := v.(type) {
	case int:
		return new(big.Float).SetInt64(int64(x)), true
	case int8:
		return new(big.Float).SetInt64(int64(x)), true
	case int16:
		return new(big.Float).SetInt64(int64(x)), true
	case int32:
		return new(big.Float).SetInt64(int64(x)), true
	case int64:
		return new(big.Float).SetInt64(x), true
	case uint:
		return new(big.Float).SetUint64(uint64(x)), true
	case uint8:
		return new(big.Float).SetUint64(uint64(x)), true
	case uint16:
		return new(big.Float).SetUint64(uint64(x)), true
	case uint32:
		return new(big.Float).SetUint64(uint64(x)), true
	case uint64:
		return new(big.Float).SetUint64(x), true
	case float32:
		return new(big.Float).SetFloat64(float64(x)), true
	case float64:
		return new(big.Float).SetFloat64(x), true
	case json.Number:
		f, ok := new(big.Float).SetString(x.String())
		return f, ok
	}
	return nil, false
}

func argBigFloat(arg string) (*big.Float, error) {
	s := argLiteral(arg)
	f, ok := new(big.Float).SetString(strings.TrimSpace(s))
	if !ok {
		return nil, fmt.Errorf("invalid number %q", arg)
	}
	return f, nil
}

func compareNumericArg(v any, arg string) (int, error) {
	vf, ok := toBigFloat(v)
	if !ok {
		return 0, fmt.Errorf("expected number, got %T", v)
	}
	af, err := argBigFloat(arg)
	if err != nil {
		return 0, err
	}
	return vf.Cmp(af), nil
}

func equalLiteral(v any, arg string) (bool, error) {
	switch x := v.(type) {
	case string:
		return x == argLiteral(arg), nil
	case bool:
		b, err := strconv.ParseBool(argLiteral(arg))
		if err != nil {
			return false, fmt.Errorf("invalid bool literal %q", arg)
		}
		return x == b, nil
	case dateValue:
		want, err := parseDateLiteral(arg)
		if err != nil {
			return false, err
		}
		return x.t.Equal(want), nil
	case nil:
		return argLiteral(arg) == "null", nil
	}

	cmp, err := compareNumericArg(v, arg)
	if err != nil {
		return false, err
	}
	return cmp == 0, nil
}

func valueLen(v any) (int, bool) {
	switch x := v.(type) {
	case string:
		return len([]rune(x)), true
	case []any:
		return len(x), true
	case map[string]any:
		return len(x), true
	}
	return 0, false
}

func canonicalString(v any) string {
	switch x := v.(type) {
	case nil:
		return ""
	case string:
		return x
	case bool:
		return strconv.FormatBool(x)
	case int:
		return strconv.FormatInt(int64(x), 10)
	case int8:
		return strconv.FormatInt(int64(x), 10)
	case int16:
		return strconv.FormatInt(int64(x), 10)
	case int32:
		return strconv.FormatInt(int64(x), 10)
	case int64:
		return strconv.FormatInt(x, 10)
	case uint:
		return strconv.FormatUint(uint64(x), 10)
	case uint8:
		return strconv.FormatUint(uint64(x), 10)
	case uint16:
		return strconv.FormatUint(uint64(x), 10)
	case uint32:
		return strconv.FormatUint(uint64(x), 10)
	case uint64:
		return strconv.FormatUint(x, 10)
	case float32:
		return strconv.FormatFloat(float64(x), 'f', -1, 64)
	case float64:
		return strconv.FormatFloat(x, 'f', -1, 64)
	case json.Number:
		return x.String()
	case dateValue:
		return x.String()
	}
	if b, err := json.Marshal(v); err == nil {
		return string(b)
	}
	return fmt.Sprint(v)
}

func parseDateLiteral(s string) (time.Time, error) {
	s = strings.TrimSpace(argLiteral(s))
	layouts := []string{
		"2006-01-02",
		"2006/01/02",
		time.RFC3339,
		"2006-01-02 15:04:05",
	}
	for _, layout := range layouts {
		if t, err := time.Parse(layout, s); err == nil {
			return t, nil
		}
	}
	return time.Time{}, fmt.Errorf("invalid date %q", s)
}

var regexCache sync.Map // map[string]*regexp.Regexp

func compileRegex(pattern string) (*regexp.Regexp, error) {
	if v, ok := regexCache.Load(pattern); ok {
		return v.(*regexp.Regexp), nil
	}
	re, err := regexp.Compile(pattern)
	if err != nil {
		return nil, fmt.Errorf("invalid regex %q: %w", pattern, err)
	}
	regexCache.Store(pattern, re)
	return re, nil
}
