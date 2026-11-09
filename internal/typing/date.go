package typing

import (
	"encoding/json"
	"strconv"
	"strings"
	"time"
)

// DateValue wraps a time.Time with the configured format so MarshalJSON
// produces the desired date string.
type DateValue struct {
	Time   time.Time
	Format string
}

func (d DateValue) MarshalJSON() ([]byte, error) {
	return json.Marshal(d.Time.Format(d.Format))
}

func (d DateValue) String() string {
	return d.Time.Format(d.Format)
}

// ConvertDate converts an Excel date serial number or a parseable date
// string into a DateValue.
func (c *Converter) ConvertDate(cell string) (any, error) {
	layout := excelDateFormatToGo(c.DateFormat)
	trimmed := strings.TrimSpace(cell)

	if serial, err := strconv.ParseInt(trimmed, 10, 64); err == nil && serial >= 1 {
		return DateValue{Time: excelSerialToTime(serial), Format: layout}, nil
	}

	if f, err := strconv.ParseFloat(trimmed, 64); err == nil && f >= 1 {
		return DateValue{Time: excelSerialToFloatTime(f), Format: layout}, nil
	}

	commonLayouts := []string{
		"2006-01-02",
		"2006/01/02",
		"01/02/2006",
		"1/2/2006",
		"2006-01-02 15:04:05",
		"2006/01/02 15:04:05",
		time.RFC3339,
	}
	for _, try := range commonLayouts {
		if t, err := time.Parse(try, trimmed); err == nil {
			return DateValue{Time: t, Format: layout}, nil
		}
	}

	return trimmed, nil
}

// excelSerialToTime converts an integer Excel date serial number to time.Time.
func excelSerialToTime(serial int64) time.Time {
	if serial < 60 {
		serial++
	}
	return time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC).AddDate(0, 0, int(serial))
}

// excelSerialToFloatTime converts a float Excel date serial number to
// time.Time, preserving the time-of-day fraction.
func excelSerialToFloatTime(serial float64) time.Time {
	whole := int64(serial)
	frac := serial - float64(whole)
	t := excelSerialToTime(whole)
	nanos := int64(frac * 86400 * 1e9)
	return t.Add(time.Duration(nanos))
}

// excelDateFormatToGo converts an Excel-style date format to a Go layout.
func excelDateFormatToGo(format string) string {
	if strings.Contains(format, "2006") || strings.Contains(format, "01") || strings.Contains(format, "02") {
		return format
	}

	var buf strings.Builder
	i := 0
	for i < len(format) {
		start := i
		for i < len(format) && isDateLetter(format[i]) {
			i++
		}
		if i > start {
			buf.WriteString(replaceDateToken(format[start:i]))
		}
		if i < len(format) {
			buf.WriteByte(format[i])
			i++
		}
	}
	return buf.String()
}

func isDateLetter(c byte) bool {
	return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')
}

func replaceDateToken(token string) string {
	switch token {
	case "yyyy":
		return "2006"
	case "yy":
		return "06"
	case "MMMM":
		return "January"
	case "MMM":
		return "Jan"
	case "MM":
		return "01"
	case "M":
		return "1"
	case "dddd":
		return "Monday"
	case "ddd":
		return "Mon"
	case "dd":
		return "02"
	case "d":
		return "2"
	case "HH", "H":
		return "15"
	case "hh":
		return "03"
	case "h":
		return "3"
	case "mm":
		return "04"
	case "m":
		return "4"
	case "ss":
		return "05"
	case "s":
		return "5"
	case "tt", "TT", "AM/PM", "am/pm":
		return "PM"
	default:
		return token
	}
}
