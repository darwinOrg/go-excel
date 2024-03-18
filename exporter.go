package dgexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"regexp"
	"strconv"
	"strings"
)

const sheetName = "Sheet1"

var columnFlags = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func getStructTagList(v any, tag string) []string {
	var resList []string
	if v == nil {
		return resList
	}

	var item any
	switch reflect.TypeOf(v).Kind() {
	case reflect.Slice, reflect.Array:
		values := reflect.ValueOf(v)
		if values.Len() == 0 {
			return resList
		}
		item = values.Index(0).Interface()
	case reflect.Struct:
		item = reflect.ValueOf(v).Interface()
	default:
		panic(fmt.Sprintf("type %v not support", reflect.TypeOf(v).Kind()))
	}

	typeOf := reflect.TypeOf(item)
	if typeOf.Kind() == reflect.Ptr {
		typeOf = typeOf.Elem()
	}

	fieldNum := typeOf.NumField()
	for i := 0; i < fieldNum; i++ {
		resList = append(resList, typeOf.Field(i).Tag.Get(tag))
	}

	return resList
}

func getTagValMap(v any, tag string) map[string]string {
	resMap := make(map[string]string)
	if v == nil {
		return resMap
	}

	isPtr := false
	typeOf := reflect.TypeOf(v)
	if typeOf.Kind() == reflect.Ptr {
		typeOf = typeOf.Elem()
		isPtr = true
	}

	fieldNum := typeOf.NumField()
	for i := 0; i < fieldNum; i++ {
		structField := typeOf.Field(i)
		tagValue := structField.Tag.Get(tag)
		rv := reflect.ValueOf(v)
		if isPtr {
			rv = rv.Elem()
		}
		val := rv.FieldByName(structField.Name)
		resMap[tagValue] = fmt.Sprintf("%v", val.Interface())
	}

	return resMap
}

func struct2MapTagList(v any, tag string) []map[string]string {
	var resList []map[string]string
	switch reflect.TypeOf(v).Kind() {
	case reflect.Slice, reflect.Array:
		values := reflect.ValueOf(v)
		for i := 0; i < values.Len(); i++ {
			resList = append(resList, getTagValMap(values.Index(i).Interface(), tag))
		}
		break
	case reflect.Struct:
		val := reflect.ValueOf(v).Interface()
		resList = append(resList, getTagValMap(val, tag))
		break
	default:
		panic(fmt.Sprintf("type %v not support", reflect.TypeOf(v).Kind()))
	}
	return resList
}

func ExportStruct2Xlsx(v any) (*excelize.File, error) {
	var tag = "excel"
	tagList := getStructTagList(v, tag)
	mapTagList := struct2MapTagList(v, tag)
	xlsx := excelize.NewFile()
	xlsx.NewSheet(sheetName)

	for c, tagVal := range tagList {
		name, _ := stringMatchExport(tagVal, regexp.MustCompile(`name\((.*?)\)`))
		cellIndex := columnFlags[c] + "1"
		xlsx.SetColWidth(sheetName, columnFlags[c], columnFlags[c], 15)
		xlsx.SetCellValue(sheetName, cellIndex, name)
	}

	for r, mapTagVal := range mapTagList {
		c := 0
		for tagKey, tagVal := range mapTagVal {
			mapping, _ := stringMatchExport(tagKey, regexp.MustCompile(`mapping\((.*?)\)`))
			if mapping != "" {
				formatStr := strings.Split(mapping, ",")
				for _, format := range formatStr {
					n := strings.SplitN(format, ":", 2)
					if len(n) != 2 {
						continue
					}
					if n[1] == tagVal {
						tagVal = n[0]
					}
				}
			}

			cellIndex := columnFlags[c] + strconv.Itoa(r+2)
			xlsx.SetCellValue(sheetName, cellIndex, tagVal)
			c++
		}
	}

	return xlsx, nil
}
