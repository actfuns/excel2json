package typing

import (
	"encoding/json"
	"math"
	"strconv"
)

// MatchJSON reports whether a decoded JSON value structurally matches t.
func MatchJSON(v any, t *Type) bool {
	if t == nil || t.Kind == KindAny {
		return true
	}
	switch t.Kind {
	case KindNull:
		return v == nil
	case KindBool:
		_, ok := v.(bool)
		return ok
	case KindString:
		_, ok := v.(string)
		return ok
	case KindInt:
		return matchInt(v, t)
	case KindUint:
		return matchUint(v, t)
	case KindFloat:
		return matchFloat(v, t)
	case KindDate:
		_, isString := v.(string)
		_, isNumber := numberValue(v)
		return isString || isNumber
	case KindSlice:
		arr, ok := v.([]any)
		if !ok {
			return false
		}
		for _, item := range arr {
			if !MatchJSON(item, t.Elem) {
				return false
			}
		}
		return true
	case KindArray:
		arr, ok := v.([]any)
		if !ok || len(arr) != t.Length {
			return false
		}
		for _, item := range arr {
			if !MatchJSON(item, t.Elem) {
				return false
			}
		}
		return true
	case KindMap:
		obj, ok := v.(map[string]any)
		if !ok {
			return false
		}
		for key, item := range obj {
			if !matchMapKey(key, t.Key) || !MatchJSON(item, t.Elem) {
				return false
			}
		}
		return true
	}
	return false
}

func matchMapKey(key string, t *Type) bool {
	if t == nil || t.Kind == KindAny || t.Kind == KindString || t.Kind == KindDate {
		return true
	}
	if t.Kind == KindInt || t.Kind == KindUint {
		return MatchJSON(json.Number(key), t)
	}
	return false
}

func numberValue(v any) (float64, bool) {
	switch x := v.(type) {
	case json.Number:
		f, err := strconv.ParseFloat(x.String(), 64)
		return f, err == nil && !math.IsNaN(f) && !math.IsInf(f, 0)
	case float64:
		return x, !math.IsNaN(x) && !math.IsInf(x, 0)
	case float32:
		return float64(x), !math.IsNaN(float64(x)) && !math.IsInf(float64(x), 0)
	case int:
		return float64(x), true
	case int8:
		return float64(x), true
	case int16:
		return float64(x), true
	case int32:
		return float64(x), true
	case int64:
		return float64(x), true
	case uint:
		return float64(x), true
	case uint8:
		return float64(x), true
	case uint16:
		return float64(x), true
	case uint32:
		return float64(x), true
	case uint64:
		return float64(x), true
	}
	return 0, false
}

func matchInt(v any, t *Type) bool {
	switch x := v.(type) {
	case json.Number:
		if _, err := strconv.ParseInt(x.String(), 10, t.Bits); err == nil {
			return true
		}
		f, err := strconv.ParseFloat(x.String(), 64)
		if err != nil || f != math.Trunc(f) {
			return false
		}
		_, err = intValue(f, t)
		return err == nil
	case int:
		return signedFits(int64(x), t.Bits)
	case int8:
		return signedFits(int64(x), t.Bits)
	case int16:
		return signedFits(int64(x), t.Bits)
	case int32:
		return signedFits(int64(x), t.Bits)
	case int64:
		return signedFits(x, t.Bits)
	case uint:
		return uint64(x) <= math.MaxInt64 && signedFits(int64(x), t.Bits)
	case uint8:
		return signedFits(int64(x), t.Bits)
	case uint16:
		return signedFits(int64(x), t.Bits)
	case uint32:
		return signedFits(int64(x), t.Bits)
	case uint64:
		return x <= math.MaxInt64 && signedFits(int64(x), t.Bits)
	case float64:
		if x != math.Trunc(x) {
			return false
		}
		_, err := intValue(x, t)
		return err == nil
	case float32:
		if float64(x) != math.Trunc(float64(x)) {
			return false
		}
		_, err := intValue(float64(x), t)
		return err == nil
	}
	return false
}

func matchUint(v any, t *Type) bool {
	switch x := v.(type) {
	case json.Number:
		if _, err := strconv.ParseUint(x.String(), 10, t.Bits); err == nil {
			return true
		}
		f, err := strconv.ParseFloat(x.String(), 64)
		if err != nil || f != math.Trunc(f) || f < 0 {
			return false
		}
		_, err = uintValue(f, t)
		return err == nil
	case uint:
		return unsignedFits(uint64(x), t.Bits)
	case uint8:
		return unsignedFits(uint64(x), t.Bits)
	case uint16:
		return unsignedFits(uint64(x), t.Bits)
	case uint32:
		return unsignedFits(uint64(x), t.Bits)
	case uint64:
		return unsignedFits(x, t.Bits)
	case int:
		if x < 0 {
			return false
		}
		return unsignedFits(uint64(x), t.Bits)
	case int8:
		if x < 0 {
			return false
		}
		return unsignedFits(uint64(x), t.Bits)
	case int16:
		if x < 0 {
			return false
		}
		return unsignedFits(uint64(x), t.Bits)
	case int32:
		if x < 0 {
			return false
		}
		return unsignedFits(uint64(x), t.Bits)
	case int64:
		if x < 0 {
			return false
		}
		return unsignedFits(uint64(x), t.Bits)
	case float64:
		if x < 0 || x != math.Trunc(x) {
			return false
		}
		_, err := uintValue(x, t)
		return err == nil
	case float32:
		if x < 0 || float64(x) != math.Trunc(float64(x)) {
			return false
		}
		_, err := uintValue(float64(x), t)
		return err == nil
	}
	return false
}

func matchFloat(v any, t *Type) bool {
	switch x := v.(type) {
	case json.Number:
		f, err := strconv.ParseFloat(x.String(), t.Bits)
		if err != nil || math.IsNaN(f) || math.IsInf(f, 0) {
			return false
		}
		return true
	case float64:
		if math.IsNaN(x) || math.IsInf(x, 0) {
			return false
		}
		return t.Bits != 32 || math.Abs(x) <= math.MaxFloat32
	case float32:
		return !math.IsNaN(float64(x)) && !math.IsInf(float64(x), 0)
	}
	if f, ok := numberValue(v); ok {
		return t.Bits != 32 || math.Abs(f) <= math.MaxFloat32
	}
	return false
}

func signedFits(v int64, bits int) bool {
	if bits >= 64 {
		return true
	}
	min := -(int64(1) << (bits - 1))
	max := (int64(1) << (bits - 1)) - 1
	return v >= min && v <= max
}

func unsignedFits(v uint64, bits int) bool {
	if bits >= 64 {
		return true
	}
	return v <= (uint64(1)<<bits)-1
}
