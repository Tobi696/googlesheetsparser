package googlesheetsparser

import (
	"errors"
	"reflect"
	"strconv"

	"google.golang.org/api/sheets/v4"
)

var Service *sheets.Service

func ParsePageIntoStructSlice[K any](spreadsheetId, sheetName string) ([]K, error) {
	resp, err := Service.Spreadsheets.Values.Get(spreadsheetId, sheetName).Do()
	if err != nil {
		return nil, err
	}
	mappings, err := CreateMappings[K](resp)
	if err != nil {
		return nil, err
	}
	var result []K
	for _, row := range resp.Values[1:] {
		var k K
		for i := range mappings {
			field := mappings[i]
			val, err := ReflectParseString(field.Type, row[i].(string))
			if err != nil {
				return nil, err
			}
			reflect.ValueOf(&k).Elem().FieldByName(field.Name).Set(val)
		}
		result = append(result, k)
	}
	return result, nil
}

func ReflectParseString(reflectType reflect.Type, cell string) (reflect.Value, error) {
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
	return reflect.ValueOf(nil), errors.New("Unsupported type: " + reflectType.Kind().String())
}

func CreateMappings[K any](data *sheets.ValueRange) (mappings []reflect.StructField, err error) {
	firstRow := data.Values[0]
	for _, cellIf := range firstRow {
		cell := cellIf.(string)
		if cell == "" {
			break
		}
		field := ReflectGetFieldByTagOrName[K](cell)
		if field == nil {
			err = errors.New("Field not found in struct: " + cell)
			return
		}
		mappings = append(mappings, *field)
	}
	return
}

func ReflectGetFieldByTagOrName[K any](name string) *reflect.StructField {
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
