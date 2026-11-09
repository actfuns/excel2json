package typing

import (
	"encoding/json"
	"math"
	"strconv"
)

// LegacyConvert applies automatic JSON/number detection when no type
// annotation is present.
func (c *Converter) LegacyConvert(cell string) (any, error) {
	if c.ParseCellJSON && cell != "" && (cell[0] == '{' || cell[0] == '[') {
		var parsed any
		if json.Unmarshal([]byte(cell), &parsed) == nil && parsed != nil {
			if c.StringifyAll {
				if b, err := json.Marshal(parsed); err == nil {
					return string(b), nil
				}
			}
			return parsed, nil
		}
	}

	if f, err := strconv.ParseFloat(cell, 64); err == nil && !math.IsNaN(f) && !math.IsInf(f, 0) {
		if i := int64(math.Round(f)); math.Abs(f-float64(i)) < 1e-9 {
			if c.StringifyAll {
				return strconv.FormatInt(i, 10), nil
			}
			return i, nil
		}
		if c.StringifyAll {
			return strconv.FormatFloat(f, 'f', -1, 64), nil
		}
		return f, nil
	}

	return cell, nil
}
