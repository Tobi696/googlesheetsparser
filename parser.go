package googlesheetsparser

import (
	"errors"
	"fmt"
	"log"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/gertd/go-pluralize"
	"google.golang.org/api/sheets/v4"
)

var pluralizeClient = pluralize.NewClient()

var (
	ErrNoSpreadSheetID       = errors.New("no spreadsheet id provided")
	ErrNoSheetName           = errors.New("no sheet name provided")
	ErrUnsupportedType       = errors.New("unsupported type")
	ErrFieldNotFoundInStruct = errors.New("field not found in struct")
	ErrInvalidDateTimeFormat = errors.New("invalid datetime format")
)

var dateTimeFormats = []string{
	"2006-01-02",
	"2006-01-02 15:04:05",
	"2006-01-02 15:04:05 -0700",
}

type Options struct {
	Service         *sheets.Service
	SpreadsheetID   string
	SheetName       string
	DatetimeFormats []string

	built bool
}

func (o Options) Build() Options {
	if o.built {
		return o
	}
	o.built = true

	o.DatetimeFormats = append(o.DatetimeFormats, dateTimeFormats...)

	return o
}

// ParseSheet parses a sheet page and returns a slice of structs with the give type.
func ParseSheetIntoStructSlice[K any](options Options) ([]K, error) {
	if !options.built {
		log.Println("googlesheetsparser: Warning: Using options that are not built")
	}

	// Set Params
	var k K
	spreadSheetId := options.SpreadsheetID
	sheetName := pluralizeClient.Plural(reflect.TypeOf(k).Name())
	if options.SheetName != "" {
		sheetName = options.SheetName
	}

	// Validate Params
	if spreadSheetId == "" {
		return nil, ErrNoSpreadSheetID
	}
	if sheetName == "" {
		return nil, ErrNoSheetName
	}

	resp, err := options.Service.Spreadsheets.Values.Get(spreadSheetId, sheetName).Do()
	if err != nil {
		return nil, err
	}

	mappings, err := createMappings[K](resp)
	if err != nil {
		return nil, err
	}

	fillEmptyValues(resp)

	var result []K
	for rowIdx, row := range resp.Values[1:] {
		var k K
		for i := range mappings {
			field := mappings[i]
			val, err := reflectParseString(field.Type, row[i].(string), options.DatetimeFormats, rowIdx, i)
			if err != nil {
				return nil, fmt.Errorf("%s: %s%d: %w", sheetName, getColumnName(i), rowIdx, err)
			}
			reflect.ValueOf(&k).Elem().FieldByName(field.Name).Set(val)
		}
		result = append(result, k)
	}
	return result, nil
}

func fillEmptyValues(data *sheets.ValueRange) {
	var maxWidth int
	for _, row := range data.Values {
		if len(row) > maxWidth {
			maxWidth = len(row)
		}
	}

	for rowIdx, row := range data.Values {
		for colIdx := len(row); colIdx < maxWidth; colIdx++ {
			data.Values[rowIdx] = append(data.Values[rowIdx], "")
		}
	}
}

func reflectParseString(pReflectType reflect.Type, cell string, dateTimeFormats []string, rowIdx, colIdx int) (reflect.Value, error) {
	reflectType := pReflectType
	var isPointer bool
	isEmpty := cell == ""
	if reflectType.Kind() == reflect.Pointer {
		reflectType = reflectType.Elem()
		isPointer = true
	}
	if isPointer && isEmpty {
		return reflect.Zero(pReflectType), nil
	}
	switch reflectType.Kind() {
	case reflect.String:
		if isPointer {
			return reflect.ValueOf(&cell), nil
		}
		return reflect.ValueOf(cell), nil
	case reflect.Int:
		i, err := strconv.Atoi(cell)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		if isPointer {
			return reflect.ValueOf(&i), nil
		}
		return reflect.ValueOf(i), nil
	case reflect.Int8:
		i, err := strconv.ParseInt(cell, 10, 8)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := int8(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Int16:
		i, err := strconv.ParseInt(cell, 10, 16)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := int16(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Int32:
		i, err := strconv.ParseInt(cell, 10, 32)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := int32(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Int64:
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := int64(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Uint:
		i, err := strconv.ParseUint(cell, 10, 0)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := uint(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Uint8:
		i, err := strconv.ParseUint(cell, 10, 8)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := uint8(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Uint16:
		i, err := strconv.ParseUint(cell, 10, 16)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := uint16(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Uint32:
		i, err := strconv.ParseUint(cell, 10, 32)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := uint32(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Uint64:
		i, err := strconv.ParseUint(cell, 10, 64)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := uint64(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Float32:
		i, err := strconv.ParseFloat(cell, 32)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := float32(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Float64:
		i, err := strconv.ParseFloat(cell, 64)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		res := float64(i)
		if isPointer {
			return reflect.ValueOf(&res), nil
		}
		return reflect.ValueOf(res), nil
	case reflect.Bool:
		i, err := strconv.ParseBool(cell)
		if err != nil {
			return reflect.ValueOf(false), err
		}
		if isPointer {
			return reflect.ValueOf(&i), nil
		}
		return reflect.ValueOf(i), nil
	case reflect.Struct:
		if reflectType.String() == "time.Time" {
			for _, dateTimeFormat := range dateTimeFormats {
				t, err := time.Parse(dateTimeFormat, cell)
				if err == nil {
					if isPointer {
						return reflect.ValueOf(&t), nil
					}
					return reflect.ValueOf(t), nil
				}
			}
			return reflect.ValueOf(time.Time{}), fmt.Errorf("%w: %s%d: %s", ErrInvalidDateTimeFormat, getColumnName(colIdx), rowIdx, cell)
		}
	}
	return reflect.ValueOf(nil), fmt.Errorf("%w: %s", ErrUnsupportedType, reflectType.Kind().String())
}

func createMappings[K any](data *sheets.ValueRange) (mappings []reflect.StructField, err error) {
	firstRow := data.Values[0]
	for colIdx, cellIf := range firstRow {
		cell := cellIf.(string)
		if cell == "" {
			break
		}
		field := reflectGetFieldByTagOrName[K](cell)
		if field == nil {
			err = fmt.Errorf("%w: %s%d: %s", ErrFieldNotFoundInStruct, getColumnName(colIdx), 1, cell)
			return
		}
		mappings = append(mappings, *field)
	}
	return
}

func reflectGetFieldByTagOrName[K any](name string) *reflect.StructField {
	var k K
	typeOfK := reflect.TypeOf(k)
	for i := 0; i < typeOfK.NumField(); i++ {
		field := typeOfK.Field(i)
		if sheetName, ok := field.Tag.Lookup("sheets"); ok && sheetName == name {
			return &field
		}
	}
	if field, ok := typeOfK.FieldByName(name); ok {
		return &field
	}
	return nil
}

func getColumnName(index int) string {
	index += 1
	var res string
	for index > 0 {
		index--
		res = string(rune(index%26+97)) + res
		index /= 26
	}
	return strings.ToUpper(res)
}
