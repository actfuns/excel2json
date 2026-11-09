package schema

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

	"github.com/actfuns/excel2json/internal/typing"
)

// Check is a parsed validator expression. Combinators (and/or/not) keep
// their operands in subs; ordinary validators keep raw literal arguments.
// Check is a parsed validator expression. Combinators keep their operands in
// Subs; ordinary validators keep raw literal arguments in Args.
type Check struct {
	Name string
	Args []string
	Subs []*Check
}

// RefResolver resolves ref(sheet.column) checks against workbook data.
type RefResolver interface {
	CheckRef(value any, sheet, column string) error
}

// Runner executes validators against values.
type Runner struct {
	Ref RefResolver
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

func checkKindMatches(kind checkKind, t *typing.Type) bool {
	if t == nil || t.Kind == typing.KindAny || kind == ckAny {
		return true
	}
	switch kind {
	case ckCollection:
		return t.Kind == typing.KindString || t.Kind == typing.KindSlice || t.Kind == typing.KindArray || t.Kind == typing.KindMap
	case ckArray:
		return t.Kind == typing.KindSlice || t.Kind == typing.KindArray
	}
	// Scalar-style checks on arrays are applied element-wise.
	if (t.Kind == typing.KindSlice || t.Kind == typing.KindArray) && t.Elem != nil {
		return checkKindMatches(kind, t.Elem)
	}
	switch kind {
	case ckNumber:
		return t.Kind == typing.KindInt || t.Kind == typing.KindUint || t.Kind == typing.KindFloat
	case ckString:
		return t.Kind == typing.KindString
	case ckScalar:
		return t.Kind == typing.KindBool || t.Kind == typing.KindString || t.Kind == typing.KindInt || t.Kind == typing.KindUint || t.Kind == typing.KindFloat || t.Kind == typing.KindDate
	case ckDate:
		return t.Kind == typing.KindDate
	}
	return false
}

func compileCheck(ch *Check, t *typing.Type) error {
	if isDirective(ch.Name) {
		return fmt.Errorf("%s may only appear directly in a rule, not inside combinators", ch.Name)
	}
	if isCombinator(ch.Name) {
		if len(ch.Subs) == 0 {
			return fmt.Errorf("%s requires at least one check", ch.Name)
		}
		for _, sub := range ch.Subs {
			if err := compileCheck(sub, t); err != nil {
				return err
			}
		}
		return nil
	}

	def, ok := checkDefs[ch.Name]
	if !ok {
		return fmt.Errorf("unknown check %q", ch.Name)
	}
	if len(ch.Args) < def.minArgs || (def.maxArgs >= 0 && len(ch.Args) > def.maxArgs) {
		if def.maxArgs < 0 {
			return fmt.Errorf("%s expects at least %d argument(s), got %d", ch.Name, def.minArgs, len(ch.Args))
		}
		if def.minArgs == def.maxArgs {
			return fmt.Errorf("%s expects %d argument(s), got %d", ch.Name, def.minArgs, len(ch.Args))
		}
		return fmt.Errorf("%s expects %d..%d argument(s), got %d", ch.Name, def.minArgs, def.maxArgs, len(ch.Args))
	}
	if !checkKindMatches(def.kind, t) {
		return fmt.Errorf("check %q does not apply to type %s", ch.Name, t)
	}
	return nil
}

func (r *Runner) RunChecks(v any, checks []*Check) error {
	arr, isArray := v.([]any)
	for _, ch := range checks {
		if isDirective(ch.Name) {
			continue
		}
		if isArray && !isCollectionCheckExpr(ch) {
			for i, item := range arr {
				if err := r.RunChecks(item, []*Check{ch}); err != nil {
					return fmt.Errorf("index %d: %w", i, err)
				}
			}
			continue
		}
		if err := r.runCheck(v, ch); err != nil {
			return err
		}
	}
	return nil
}

// isCollectionCheckExpr reports whether a check (or any operand of a
// combinator) targets the collection as a whole rather than its elements.
func HasPad(checks []*Check) bool {
	for _, ch := range checks {
		if ch.Name == "pad" {
			return true
		}
	}
	return false
}

func isCollectionCheckExpr(ch *Check) bool {
	if isCombinator(ch.Name) {
		for _, sub := range ch.Subs {
			if isCollectionCheckExpr(sub) {
				return true
			}
		}
		return false
	}
	switch ch.Name {
	case "len", "minlen", "maxlen", "unique", "nonempty":
		return true
	}
	return false
}

func (r *Runner) runCheck(v any, ch *Check) error {
	switch ch.Name {
	case "and":
		for _, sub := range ch.Subs {
			if err := r.runCheck(v, sub); err != nil {
				return err
			}
		}
		return nil
	case "or":
		var errs []string
		for _, sub := range ch.Subs {
			if err := r.runCheck(v, sub); err == nil {
				return nil
			} else {
				errs = append(errs, err.Error())
			}
		}
		return fmt.Errorf("or: none matched (%s)", strings.Join(errs, "; "))
	case "not":
		allPass := true
		for _, sub := range ch.Subs {
			if err := r.runCheck(v, sub); err != nil {
				allPass = false
			}
		}
		if allPass {
			return fmt.Errorf("not: condition matched")
		}
		return nil
	}

	def := checkDefs[ch.Name]
	if def.kind == ckNumber || def.kind == ckScalar {
		if _, ok := toBigFloat(v); !ok {
			if def.kind == ckNumber {
				return fmt.Errorf("%s expects a number, got %T", ch.Name, v)
			}
		}
	}

	switch ch.Name {
	case "gt", "gte", "lt", "lte":
		cmp, err := compareNumericArg(v, ch.Args[0])
		if err != nil {
			return fmt.Errorf("%s: %w", ch.Name, err)
		}
		ok := cmp > 0
		if ch.Name == "gte" {
			ok = cmp >= 0
		} else if ch.Name == "lt" {
			ok = cmp < 0
		} else if ch.Name == "lte" {
			ok = cmp <= 0
		}
		if !ok {
			return fmt.Errorf("%s(%s) failed for %s", ch.Name, ch.Args[0], CanonicalString(v))
		}
		return nil

	case "between":
		lo, err := compareNumericArg(v, ch.Args[0])
		if err != nil {
			return fmt.Errorf("between: %w", err)
		}
		hi, err := compareNumericArg(v, ch.Args[1])
		if err != nil {
			return fmt.Errorf("between: %w", err)
		}
		if lo < 0 || hi > 0 {
			return fmt.Errorf("between(%s,%s) failed for %s", ch.Args[0], ch.Args[1], CanonicalString(v))
		}
		return nil

	case "eq", "ne":
		equal, err := equalLiteral(v, ch.Args[0])
		if err != nil {
			return fmt.Errorf("%s: %w", ch.Name, err)
		}
		if (ch.Name == "eq") != equal {
			return fmt.Errorf("%s(%s) failed for %s", ch.Name, ch.Args[0], CanonicalString(v))
		}
		return nil

	case "positive", "negative", "nonzero":
		cmp, err := compareNumericArg(v, "0")
		if err != nil {
			return fmt.Errorf("%s: %w", ch.Name, err)
		}
		ok := cmp != 0
		if ch.Name == "positive" {
			ok = cmp > 0
		} else if ch.Name == "negative" {
			ok = cmp < 0
		}
		if !ok {
			return fmt.Errorf("%s failed for %s", ch.Name, CanonicalString(v))
		}
		return nil

	case "multipleof":
		vf, ok := toBigRat(v)
		if !ok {
			return fmt.Errorf("multipleof expects a number, got %T", v)
		}
		af, err := argBigRat(ch.Args[0])
		if err != nil {
			return fmt.Errorf("multipleof: %w", err)
		}
		if af.Sign() == 0 {
			return fmt.Errorf("multipleof divisor cannot be zero")
		}
		q := new(big.Rat).Quo(vf, af)
		if !q.IsInt() {
			return fmt.Errorf("multipleof(%s) failed for %s", ch.Args[0], CanonicalString(v))
		}
		return nil

	case "len", "minlen", "maxlen":
		n, ok := valueLen(v)
		if !ok {
			return fmt.Errorf("%s expects string/array/map, got %T", ch.Name, v)
		}
		switch ch.Name {
		case "len":
			if len(ch.Args) == 1 {
				want, err := strconv.Atoi(ch.Args[0])
				if err != nil {
					return fmt.Errorf("len: invalid length %q", ch.Args[0])
				}
				if n != want {
					return fmt.Errorf("len(%d) failed, got %d", want, n)
				}
				return nil
			}
			lo, err := strconv.Atoi(ch.Args[0])
			if err != nil {
				return fmt.Errorf("len: invalid length %q", ch.Args[0])
			}
			hi, err := strconv.Atoi(ch.Args[1])
			if err != nil {
				return fmt.Errorf("len: invalid length %q", ch.Args[1])
			}
			if n < lo || n > hi {
				return fmt.Errorf("len(%d,%d) failed, got %d", lo, hi, n)
			}
			return nil
		case "minlen":
			lo, err := strconv.Atoi(ch.Args[0])
			if err != nil {
				return fmt.Errorf("minlen: invalid length %q", ch.Args[0])
			}
			if n < lo {
				return fmt.Errorf("minlen(%d) failed, got %d", lo, n)
			}
			return nil
		default:
			hi, err := strconv.Atoi(ch.Args[0])
			if err != nil {
				return fmt.Errorf("maxlen: invalid length %q", ch.Args[0])
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
		if len(ch.Args) == 1 {
			switch ArgLiteral(ch.Args[0]) {
			case "", "cn":
				// keep default
			case "e164":
				pattern = `^\+?[1-9]\d{6,14}$`
			default:
				pattern = ArgLiteral(ch.Args[0])
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
		re, err := compileRegex(ArgLiteral(ch.Args[0]))
		if err != nil {
			return fmt.Errorf("regex: %w", err)
		}
		if !re.MatchString(s) {
			return fmt.Errorf("regex(%s) failed for %q", ch.Args[0], s)
		}
		return nil

	case "prefix", "suffix", "contains":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("%s expects string, got %T", ch.Name, v)
		}
		arg := ArgLiteral(ch.Args[0])
		var okVal bool
		switch ch.Name {
		case "prefix":
			okVal = strings.HasPrefix(s, arg)
		case "suffix":
			okVal = strings.HasSuffix(s, arg)
		default:
			okVal = strings.Contains(s, arg)
		}
		if !okVal {
			return fmt.Errorf("%s(%q) failed for %q", ch.Name, arg, s)
		}
		return nil

	case "alpha", "alnum", "numeric":
		s, ok := v.(string)
		if !ok {
			return fmt.Errorf("%s expects string, got %T", ch.Name, v)
		}
		for _, r := range s {
			switch ch.Name {
			case "alpha":
				ok = unicode.IsLetter(r)
			case "alnum":
				ok = unicode.IsLetter(r) || unicode.IsDigit(r)
			default:
				ok = unicode.IsDigit(r)
			}
			if !ok {
				return fmt.Errorf("%s failed for %q", ch.Name, s)
			}
		}
		return nil

	case "oneof", "in", "notoneof", "notin":
		any := false
		for _, arg := range ch.Args {
			equal, err := equalLiteral(v, arg)
			if err != nil {
				return fmt.Errorf("%s: %w", ch.Name, err)
			}
			if equal {
				any = true
				break
			}
		}
		if ch.Name == "oneof" || ch.Name == "in" {
			if !any {
				return fmt.Errorf("%s(%s) failed for %s", ch.Name, strings.Join(ch.Args, ","), CanonicalString(v))
			}
			return nil
		}
		if any {
			return fmt.Errorf("%s(%s) failed for %s", ch.Name, strings.Join(ch.Args, ","), CanonicalString(v))
		}
		return nil

	case "unique":
		arr, ok := v.([]any)
		if !ok {
			return fmt.Errorf("unique expects array, got %T", v)
		}
		seen := make(map[string]struct{}, len(arr))
		for _, item := range arr {
			k := CanonicalString(item)
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
		return r.runRefCheck(v, ch.Args[0])

	case "after", "before":
		dv, ok := v.(typing.DateValue)
		if !ok {
			return fmt.Errorf("%s expects date, got %T", ch.Name, v)
		}
		want, err := parseDateLiteral(ch.Args[0])
		if err != nil {
			return fmt.Errorf("%s: %w", ch.Name, err)
		}
		if ch.Name == "after" {
			if !dv.Time.After(want) {
				return fmt.Errorf("after(%s) failed for %s", ch.Args[0], dv.String())
			}
		} else if !dv.Time.Before(want) {
			return fmt.Errorf("before(%s) failed for %s", ch.Args[0], dv.String())
		}
		return nil
	}

	return fmt.Errorf("unknown check %q", ch.Name)
}

func ArgLiteral(s string) string {
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
		v, err := unquoteSingle(s)
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
		if f, ok := new(big.Float).SetString(strconv.FormatFloat(float64(x), 'g', -1, 32)); ok {
			return f, true
		}
		return new(big.Float).SetFloat64(float64(x)), true
	case float64:
		if f, ok := new(big.Float).SetString(strconv.FormatFloat(x, 'g', -1, 64)); ok {
			return f, true
		}
		return new(big.Float).SetFloat64(x), true
	case json.Number:
		f, ok := new(big.Float).SetString(x.String())
		return f, ok
	}
	return nil, false
}

func toBigRat(v any) (*big.Rat, bool) {
	switch x := v.(type) {
	case int:
		return new(big.Rat).SetInt64(int64(x)), true
	case int8:
		return new(big.Rat).SetInt64(int64(x)), true
	case int16:
		return new(big.Rat).SetInt64(int64(x)), true
	case int32:
		return new(big.Rat).SetInt64(int64(x)), true
	case int64:
		return new(big.Rat).SetInt64(x), true
	case uint:
		return new(big.Rat).SetUint64(uint64(x)), true
	case uint8:
		return new(big.Rat).SetUint64(uint64(x)), true
	case uint16:
		return new(big.Rat).SetUint64(uint64(x)), true
	case uint32:
		return new(big.Rat).SetUint64(uint64(x)), true
	case uint64:
		return new(big.Rat).SetUint64(x), true
	case float32:
		return new(big.Rat).SetString(strconv.FormatFloat(float64(x), 'g', -1, 32))
	case float64:
		return new(big.Rat).SetString(strconv.FormatFloat(x, 'g', -1, 64))
	case json.Number:
		return new(big.Rat).SetString(x.String())
	}
	return nil, false
}

func argBigRat(arg string) (*big.Rat, error) {
	s := ArgLiteral(arg)
	r, ok := new(big.Rat).SetString(strings.TrimSpace(s))
	if !ok {
		return nil, fmt.Errorf("invalid number %q", arg)
	}
	return r, nil
}

func argBigFloat(arg string) (*big.Float, error) {
	s := ArgLiteral(arg)
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
		return x == ArgLiteral(arg), nil
	case bool:
		b, err := strconv.ParseBool(ArgLiteral(arg))
		if err != nil {
			return false, fmt.Errorf("invalid bool literal %q", arg)
		}
		return x == b, nil
	case typing.DateValue:
		want, err := parseDateLiteral(arg)
		if err != nil {
			return false, err
		}
		return x.Time.Equal(want), nil
	case nil:
		return ArgLiteral(arg) == "null", nil
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

func CanonicalString(v any) string {
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
	case typing.DateValue:
		return x.String()
	}
	if b, err := json.Marshal(v); err == nil {
		return string(b)
	}
	return fmt.Sprint(v)
}

func parseDateLiteral(s string) (time.Time, error) {
	s = strings.TrimSpace(ArgLiteral(s))
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

func compileChecks(checks []*Check, t *typing.Type) error {
	padSeen := false
	for _, ch := range checks {
		if isDirective(ch.Name) {
			if padSeen {
				return fmt.Errorf("duplicate pad directive")
			}
			padSeen = true
			def, ok := checkDefs[ch.Name]
			if !ok {
				return fmt.Errorf("unknown check %q", ch.Name)
			}
			if len(ch.Args) > def.maxArgs {
				return fmt.Errorf("pad expects at most %d argument(s), got %d", def.maxArgs, len(ch.Args))
			}
			if t == nil || t.Kind != typing.KindArray {
				return fmt.Errorf("pad requires a fixed-size array type")
			}
			continue
		}
		if err := compileCheck(ch, t); err != nil {
			return err
		}
	}
	return nil
}

func (r *Runner) runRefCheck(v any, arg string) error {
	arg = strings.TrimSpace(arg)
	dot := strings.LastIndexByte(arg, '.')
	if dot <= 0 || dot == len(arg)-1 {
		return fmt.Errorf("ref expects package.column, got %q", arg)
	}
	sheet, column := arg[:dot], arg[dot+1:]
	if r.Ref == nil {
		return fmt.Errorf("ref(%s.%s) requires workbook context", sheet, column)
	}
	return r.Ref.CheckRef(v, sheet, column)
}

func parseChecks(text string) ([]*Check, error) {
	text = strings.TrimSpace(text)
	if text == "" {
		return nil, nil
	}
	parts, err := splitTopLevel(text, ',')
	if err != nil {
		return nil, fmt.Errorf("invalid checks %q: %w", text, err)
	}
	var checks []*Check
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

func parseCheck(text string) (*Check, error) {
	text = strings.TrimSpace(text)
	i := 0
	for i < len(text) && isCheckNameChar(text[i]) {
		i++
	}
	if i == 0 {
		return nil, fmt.Errorf("invalid check %q", text)
	}

	ch := &Check{Name: strings.ToLower(text[:i])}
	rest := text[i:]
	if rest == "" {
		return ch, nil
	}
	if rest[0] != '(' {
		return nil, fmt.Errorf("invalid check %q: expected '(' after %s", text, ch.Name)
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
			ch.Args = append(ch.Args, arg)
		}
	}

	if isCombinator(ch.Name) {
		for _, arg := range ch.Args {
			sub, err := parseCheck(arg)
			if err != nil {
				return nil, fmt.Errorf("invalid %s operand %q: %w", ch.Name, arg, err)
			}
			ch.Subs = append(ch.Subs, sub)
		}
		ch.Args = nil
	}
	return ch, nil
}

func isCheckNameChar(c byte) bool {
	return c == '_' ||
		(c >= 'a' && c <= 'z') ||
		(c >= 'A' && c <= 'Z') ||
		(c >= '0' && c <= '9')
}

func unquoteSingle(s string) (string, error) {
	if len(s) < 2 || s[len(s)-1] != s[0] {
		return "", fmt.Errorf("invalid quoted string %q", s)
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
			return "", fmt.Errorf("invalid escape in %q", s)
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
				return "", fmt.Errorf("invalid unicode escape in %q", s)
			}
			r, err := strconv.ParseUint(s[i+1:i+5], 16, 16)
			if err != nil {
				return "", fmt.Errorf("invalid unicode escape in %q: %w", s, err)
			}
			b.WriteRune(rune(r))
			i += 4
		default:
			return "", fmt.Errorf("invalid escape %q", s[i])
		}
	}
	return b.String(), nil
}
