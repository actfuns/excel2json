package main

import (
	"encoding/json"
	"testing"
)

func TestParseType(t *testing.T) {
	valid := []struct {
		in   string
		want string
	}{
		{"int", "int"},
		{"int8", "int8"},
		{"int16", "int16"},
		{"int32", "int32"},
		{"int64", "int64"},
		{"uint", "uint"},
		{"uint8", "uint8"},
		{"uint16", "uint16"},
		{"uint32", "uint32"},
		{"uint64", "uint64"},
		{"uintptr", "uintptr"},
		{"float", "float"},
		{"float32", "float32"},
		{"float64", "float64"},
		{"bool", "bool"},
		{"string", "string"},
		{"byte", "byte"},
		{"rune", "rune"},
		{"date", "date"},
		{"object", "map[string]any"},
		{"list", "[]any"},
		{"interface{}", "any"},
		{"[]int", "[]int"},
		{"[3]int", "[3]int"},
		{"map[string]string", "map[string]string"},
		{"map[string][]int", "map[string][]int"},
		{"[2][2]int", "[2][2]int"},
		{"int!", "int!"},
		{"[3]int?", "[3]int?"},
	}
	for _, c := range valid {
		got, err := parseType(c.in)
		if err != nil {
			t.Errorf("parseType(%q): %v", c.in, err)
			continue
		}
		if got.String() != c.want {
			t.Errorf("parseType(%q) = %q, want %q", c.in, got.String(), c.want)
		}
	}

	var nilType *goType
	if nilType.String() != "<nil>" {
		t.Errorf("nil type String = %q", nilType.String())
	}

	invalid := []string{"", "?", "!", "[]", "[3]", "map[string]", "map[string][]", "map[string", "int[", "[x]int", "intt", "map[string]interesting", "map[string]int["}
	for _, in := range invalid {
		if _, err := parseType(in); err == nil {
			t.Errorf("parseType(%q): expected error", in)
		}
	}
}

func TestConvertTyped(t *testing.T) {
	e := NewExporter(Options{})
	valid := []struct {
		typ, cell, want string
	}{
		{"int", "42", "42"},
		{"int8", "-128", "-128"},
		{"uint32", "4294967295", "4294967295"},
		{"float32", "1.5", "1.5"},
		{"bool", "true", "true"},
		{"string", "hi", `"hi"`},
		{"[]int", "[1,2,3]", "[1,2,3]"},
		{"[3]int", "[4,5,6]", "[4,5,6]"},
		{"[2][2]int", "[[1,2],[3,4]]", "[[1,2],[3,4]]"},
		{"[3]int64", "[9223372036854775807,-9223372036854775808,0]", "[9223372036854775807,-9223372036854775808,0]"},
		{"[2]uint64", "[18446744073709551615,0]", "[18446744073709551615,0]"},
		{"[2]uint8", "[1.0,255.0]", "[1,255]"},
		{"[3]int", "[1.0,2.0,3.0]", "[1,2,3]"},
		{"[1]int", "[1e3]", "[1000]"},
		{"[0]int", "[]", "[]"},
		{"[]int", "[]", "[]"},
		{"map[string]string", "{}", "{}"},
		{"map[string]string", `{"a":"b"}`, `{"a":"b"}`},
		{"map[string][]int", `{"a":[1,2]}`, `{"a":[1,2]}`},
		{"any", "123", "123"},
		{"object", `{"x":1}`, `{"x":1}`},
		{"list", `[1,"a",true]`, `[1,"a",true]`},
	}
	for _, c := range valid {
		v, err := e.convertTyped(c.typ, c.cell)
		if err != nil {
			t.Errorf("%s %q: %v", c.typ, c.cell, err)
			continue
		}
		b, _ := json.Marshal(v)
		if string(b) != c.want {
			t.Errorf("%s %q = %s, want %s", c.typ, c.cell, b, c.want)
		}
	}
}

func TestEmptyAndRequired(t *testing.T) {
	cases := []struct {
		opts Options
		typ  string
		want string
		err  bool
	}{
		{Options{}, "int", "0", false},
		{Options{}, "bool", "false", false},
		{Options{}, "string", `""`, false},
		{Options{}, "[]int", "null", false},
		{Options{}, "map[string]int", "null", false},
		{Options{}, "int!", "", true},
		{Options{}, "int?", "0", false},
		{Options{Required: true}, "int", "", true},
		{Options{Required: true}, "int?", "0", false},
		{Options{Required: true}, "int!", "", true},
	}
	for _, c := range cases {
		v, err := NewExporter(c.opts).convertTyped(c.typ, "")
		if (err != nil) != c.err {
			t.Errorf("%s required=%v: err=%v, wantErr=%v", c.typ, c.opts.Required, err, c.err)
			continue
		}
		if err != nil {
			continue
		}
		b, _ := json.Marshal(v)
		if string(b) != c.want {
			t.Errorf("%s required=%v = %s, want %s", c.typ, c.opts.Required, b, c.want)
		}
	}
}

func TestInvalidValues(t *testing.T) {
	e := NewExporter(Options{})
	cases := []struct{ typ, cell string }{
		{"int", "abc"},
		{"int8", "128"},
		{"uint", "-1"},
		{"uint8", "256"},
		{"float32", "1e40"},
		{"bool", "maybe"},
		{"float", "NaN"},
		{"float", "+Inf"},
		{"[]int", `[1,"x"]`},
		{"[3]int", "[1,2]"},
		{"[1]uint8", "[256.0]"},
		{"[1]int8", "[128.0]"},
		{"[1]int", "[1e400]"},
		{"[1]float32", "[1e40]"},
		{"[]int", "[1] [2]"},
		{"object", `{"a":1} {"b":2}`},
		{"map[string]string", `{"a":1}`},
		{"map[int]string", `{"a":"b"}`},
		{"object", "[1,2]"},
		{"list", `{"a":1}`},
	}
	for _, c := range cases {
		if _, err := e.convertTyped(c.typ, c.cell); err == nil {
			t.Errorf("%s %q: expected error", c.typ, c.cell)
		}
	}
}

func TestDateConversion(t *testing.T) {
	e := NewExporter(Options{DateFormat: "yyyy/MM/dd"})
	if v, err := e.convertTyped("map[date]string", `{"2024-01-02":"x"}`); err != nil {
		t.Errorf("map[date]string: %v", err)
	} else if b, _ := json.Marshal(v); string(b) != `{"2024/01/02":"x"}` {
		t.Errorf("map[date]string = %s", b)
	}
	cases := []struct{ cell, want string }{
		{"2024-01-02", `"2024/01/02"`},
		{"2024/01/02", `"2024/01/02"`},
		{"not-a-date", `"not-a-date"`},
	}
	for _, c := range cases {
		v, err := e.convertTyped("date", c.cell)
		if err != nil {
			t.Errorf("date %q: %v", c.cell, err)
			continue
		}
		b, _ := json.Marshal(v)
		if string(b) != c.want {
			t.Errorf("date %q = %s, want %s", c.cell, b, c.want)
		}
	}
}

func TestIntegerBoundaries(t *testing.T) {
	e := NewExporter(Options{})
	valid := []struct{ typ, cell string }{
		{"int", "-9223372036854775808"}, {"int", "9223372036854775807"},
		{"int8", "-128"}, {"int8", "127"},
		{"int16", "-32768"}, {"int16", "32767"},
		{"int32", "-2147483648"}, {"int32", "2147483647"},
		{"int64", "-9223372036854775808"}, {"int64", "9223372036854775807"},
		{"uint", "0"}, {"uint", "18446744073709551615"},
		{"uint8", "0"}, {"uint8", "255"},
		{"uint16", "0"}, {"uint16", "65535"},
		{"uint32", "0"}, {"uint32", "4294967295"},
		{"uint64", "0"}, {"uint64", "18446744073709551615"},
		{"uintptr", "0"}, {"uintptr", "18446744073709551615"},
	}
	for _, c := range valid {
		if _, err := e.convertTyped(c.typ, c.cell); err != nil {
			t.Errorf("%s %s: unexpected error: %v", c.typ, c.cell, err)
		}
	}

	invalid := []struct{ typ, cell string }{
		{"int8", "-129"}, {"int8", "128"},
		{"int16", "-32769"}, {"int16", "32768"},
		{"int32", "-2147483649"}, {"int32", "2147483648"},
		{"int64", "-9223372036854775809"}, {"int64", "9223372036854775808"},
		{"uint8", "-1"}, {"uint8", "256"},
		{"uint16", "-1"}, {"uint16", "65536"},
		{"uint32", "-1"}, {"uint32", "4294967296"},
		{"uint64", "-1"}, {"uint64", "18446744073709551616"},
	}
	for _, c := range invalid {
		if _, err := e.convertTyped(c.typ, c.cell); err == nil {
			t.Errorf("%s %s: expected error", c.typ, c.cell)
		}
	}
}

func TestExcelSerialTime(t *testing.T) {
	cases := []struct {
		serial int64
		want   string
	}{
		{1, "1900-01-01"},
		{59, "1900-02-28"},
		{60, "1900-02-28"},
		{61, "1900-03-01"},
		{45293, "2024-01-02"},
	}
	for _, c := range cases {
		got := excelSerialToTime(c.serial).Format("2006-01-02")
		if got != c.want {
			t.Errorf("excelSerialToTime(%d) = %s, want %s", c.serial, got, c.want)
		}
	}

	if got := excelSerialToFloatTime(45293.5).Format("2006-01-02 15:04:05"); got != "2024-01-02 12:00:00" {
		t.Errorf("excelSerialToFloatTime(45293.5) = %s", got)
	}
}

func TestExcelDateFormatToGo(t *testing.T) {
	cases := []struct{ in, want string }{
		{"yyyy/MM/dd", "2006/01/02"},
		{"yyyy-MM-dd HH:mm:ss", "2006-01-02 15:04:05"},
		{"MM/dd/yyyy", "01/02/2006"},
		{"yyyy年MM月dd日", "2006年01月02日"},
		{"2006-01-02", "2006-01-02"},
	}
	for _, c := range cases {
		if got := excelDateFormatToGo(c.in); got != c.want {
			t.Errorf("excelDateFormatToGo(%q) = %q, want %q", c.in, got, c.want)
		}
	}
}

func TestZeroValue(t *testing.T) {
	cases := []struct{ typ, want string }{
		{"bool", "false"},
		{"string", `""`},
		{"date", `""`},
		{"int", "0"},
		{"uint", "0"},
		{"float", "0"},
		{"any", "null"},
		{"object", "null"},
		{"list", "null"},
		{"[]int", "null"},
		{"[2]int", "null"},
		{"map[string]int", "null"},
	}
	for _, c := range cases {
		typ, err := parseType(c.typ)
		if err != nil {
			t.Fatal(err)
		}
		b, _ := json.Marshal(zeroValue(typ))
		if string(b) != c.want {
			t.Errorf("zeroValue(%s) = %s, want %s", c.typ, b, c.want)
		}
	}
}

func TestJsonNumber(t *testing.T) {
	cases := []struct {
		in   any
		want float64
		ok   bool
	}{
		{float64(1.5), 1.5, true},
		{float32(2.5), 2.5, true},
		{int(3), 3, true},
		{int8(4), 4, true},
		{int16(5), 5, true},
		{int32(6), 6, true},
		{int64(7), 7, true},
		{uint(8), 8, true},
		{uint8(9), 9, true},
		{uint16(10), 10, true},
		{uint32(11), 11, true},
		{uint64(12), 12, true},
		{json.Number("5.5"), 5.5, true},
		{json.Number("bad"), 0, false},
		{"x", 0, false},
		{nil, 0, false},
	}
	for _, c := range cases {
		got, ok := jsonNumber(c.in)
		if ok != c.ok || (ok && got != c.want) {
			t.Errorf("jsonNumber(%v) = %v, %v; want %v, %v", c.in, got, ok, c.want, c.ok)
		}
	}
}

func TestReplaceDateToken(t *testing.T) {
	cases := map[string]string{
		"yyyy": "2006", "yy": "06",
		"MMMM": "January", "MMM": "Jan", "MM": "01", "M": "1",
		"dddd": "Monday", "ddd": "Mon", "dd": "02", "d": "2",
		"HH": "15", "H": "15", "hh": "03", "h": "3",
		"mm": "04", "m": "4", "ss": "05", "s": "5",
		"tt": "PM", "TT": "PM", "AM/PM": "PM", "am/pm": "PM",
		"unknown": "unknown",
	}
	for in, want := range cases {
		if got := replaceDateToken(in); got != want {
			t.Errorf("replaceDateToken(%q) = %q, want %q", in, got, want)
		}
	}
}

func TestConvertCellBranches(t *testing.T) {
	e := NewExporter(Options{})
	cases := []struct {
		colType, cell, want string
	}{
		{"", "", `""`},
		{"", "12.0", "12"},
		{"int", "1", "1"},
		{"string", "x", `"x"`},
	}
	for _, c := range cases {
		v, err := e.convertCell(c.colType, c.cell)
		if err != nil {
			t.Errorf("convertCell(%q, %q): %v", c.colType, c.cell, err)
			continue
		}
		b, _ := json.Marshal(v)
		if string(b) != c.want {
			t.Errorf("convertCell(%q, %q) = %s, want %s", c.colType, c.cell, b, c.want)
		}
	}
}

func TestConvertJSONDirect(t *testing.T) {
	e := NewExporter(Options{DateFormat: "yyyy-MM-dd"})
	cases := []struct {
		typ  string
		v    any
		want string
	}{
		{"any", int64(1), "1"},
		{"int", int64(5), "5"},
		{"int", json.Number("6"), "6"},
		{"int", float64(7), "7"},
		{"uint", uint64(8), "8"},
		{"uint", json.Number("9"), "9"},
		{"float", float32(1.5), "1.5"},
		{"float", int64(10), "10"},
		{"date", json.Number("45293"), `"2024-01-02"`},
		{"date", int64(1), `"1900-01-01"`},
		{"[]int", []any{int64(1), json.Number("2")}, "[1,2]"},
	}
	for _, c := range cases {
		typ, err := parseType(c.typ)
		if err != nil {
			t.Fatal(err)
		}
		v, err := e.convertJSON(c.v, typ)
		if err != nil {
			t.Errorf("%s %v: %v", c.typ, c.v, err)
			continue
		}
		b, _ := json.Marshal(v)
		if string(b) != c.want {
			t.Errorf("%s %v = %s, want %s", c.typ, c.v, b, c.want)
		}
	}
}
