package googlesheetsparser

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"

	"github.com/gertd/go-pluralize"
	"google.golang.org/api/sheets/v4"
)

var Service *sheets.Service
var SpreadSheetID string

var pluralizeClient = pluralize.NewClient()

var (
	ErrNoSpreadSheetID       = errors.New("no spreadsheet id provided")
	ErrNoSheetName           = errors.New("no sheet name provided")
	ErrUnsupportedType       = errors.New("u‚nsupported type")
	ErrFieldNotFoundInStruct = errors.New("field not found in struct")
)

type Options struct {
	SpreadSheetID string
	SheetName     string
}

// ParseSheet parses a sheet page and returns a slice of structs with the give type.
func ParsePageIntoStructSlice[K any](options *Options) ([]K, error) {
	// Set Params
	var k K
	spreadSheetId := SpreadSheetID
	sheetName := pluralizeClient.Plural(reflect.TypeOf(k).Name())
	if options != nil {
		if options.SpreadSheetID != "" {
			spreadSheetId = options.SpreadSheetID
		}
		if options.SheetName != "" {
			sheetName = options.SheetName
		}
	}

	// Validate Params
	if spreadSheetId == "" {
		return nil, ErrNoSpreadSheetID
	}
	if sheetName == "" {
		return nil, ErrNoSheetName
	}

	resp, err := Service.Spreadsheets.Values.Get(spreadSheetId, sheetName).Do()
	if err != nil {
		return nil, err
	}
	mappings, err := createMappings[K](resp)
	if err != nil {
		return nil, err
	}
	var result []K
	for _, row := range resp.Values[1:] {
		var k K
		for i := range mappings {
			field := mappings[i]
			val, err := reflectParseString(field.Type, row[i].(string))
			if err != nil {
				return nil, err
			}
			reflect.ValueOf(&k).Elem().FieldByName(field.Name).Set(val)
		}
		result = append(result, k)
	}
	return result, nil
}

func reflectParseString(reflectType reflect.Type, cell string) (reflect.Value, error) {
	switch reflectType.Kind() {
	case reflect.String:
		return reflect.ValueOf(cell), nil
	case reflect.Int:
		i, err := strconv.Atoi(cell)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(i), nil
	case reflect.Int8:
		i, err := strconv.ParseInt(cell, 10, 8)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(int8(i)), nil
	case reflect.Int16:
		i, err := strconv.ParseInt(cell, 10, 16)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(int16(i)), nil
	case reflect.Int32:
		i, err := strconv.ParseInt(cell, 10, 32)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(int32(i)), nil
	case reflect.Int64:
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(i), nil
	case reflect.Uint:
		i, err := strconv.ParseUint(cell, 10, 0)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(uint(i)), nil
	case reflect.Uint8:
		i, err := strconv.ParseUint(cell, 10, 8)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(uint8(i)), nil
	case reflect.Uint16:
		i, err := strconv.ParseUint(cell, 10, 16)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(uint16(i)), nil
	case reflect.Uint32:
		i, err := strconv.ParseUint(cell, 10, 32)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(uint32(i)), nil
	case reflect.Uint64:
		i, err := strconv.ParseUint(cell, 10, 64)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(i), nil
	case reflect.Float32:
		i, err := strconv.ParseFloat(cell, 32)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(float32(i)), nil
	case reflect.Float64:
		i, err := strconv.ParseFloat(cell, 64)
		if err != nil {
			return reflect.ValueOf(0), err
		}
		return reflect.ValueOf(i), nil
	case reflect.Bool:
		i, err := strconv.ParseBool(cell)
		if err != nil {
			return reflect.ValueOf(false), err
		}
		return reflect.ValueOf(i), nil
	}
	return reflect.ValueOf(nil), fmt.Errorf("%w: %s", ErrUnsupportedType, reflectType.Kind().String())
}

func createMappings[K any](data *sheets.ValueRange) (mappings []reflect.StructField, err error) {
	firstRow := data.Values[0]
	for _, cellIf := range firstRow {
		cell := cellIf.(string)
		if cell == "" {
			break
		}
		field := reflectGetFieldByTagOrName[K](cell)
		if field == nil {
			err = fmt.Errorf("%w: %s", ErrFieldNotFoundInStruct, cell)
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
