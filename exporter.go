package dgexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"regexp"
	"strconv"
	"strings"
)

const defaultSheetName = "Sheet1"

var (
	columnFlags = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	urlRegex    = regexp.MustCompile(`^((https|http|ftp|rtsp|mms)?://)\S+$`)
)

type ExcelHeader struct {
	Name  string
	Width float64
}

type ExcelSheet struct {
	Name    string
	Headers []*ExcelHeader
	Datas   [][]any
}

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

func getTagValMap(v any, tag string) []string {
	if v == nil {
		return []string{}
	}

	isPtr := false
	typeOf := reflect.TypeOf(v)
	if typeOf.Kind() == reflect.Ptr {
		typeOf = typeOf.Elem()
		isPtr = true
	}

	var resMap []string
	fieldNum := typeOf.NumField()
	for i := 0; i < fieldNum; i++ {
		structField := typeOf.Field(i)
		//tagValue := structField.Tag.Get(tag)
		rv := reflect.ValueOf(v)
		if isPtr {
			rv = rv.Elem()
		}
		val := rv.FieldByName(structField.Name)
		resMap = append(resMap, fmt.Sprintf("%v", val.Interface()))
	}

	return resMap
}

func struct2MapTagList(v any, tag string) [][]string {
	var resList [][]string
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
	_, _ = xlsx.NewSheet(defaultSheetName)

	for c, tagVal := range tagList {
		name, _ := stringMatchExport(tagVal, regexp.MustCompile(`name\((.*?)\)`))

		width, _ := stringMatchExport(tagVal, regexp.MustCompile(`width\((.*?)\)`))
		if width == "" {
			width = "20"
		}
		wt, _ := strconv.Atoi(width)
		if wt == 0 {
			wt = 20
		}
		_ = xlsx.SetColWidth(defaultSheetName, columnFlags[c], columnFlags[c], float64(wt))

		cellIndex := columnFlags[c] + "1"
		_ = xlsx.SetCellValue(defaultSheetName, cellIndex, name)
	}

	for r, mapTagVal := range mapTagList {
		c := 0
		for i, tagVal := range mapTagVal {
			tagKey := tagList[i]
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
			if urlRegex.MatchString(tagVal) {
				_ = xlsx.SetCellFormula(defaultSheetName, cellIndex, fmt.Sprintf("=HYPERLINK(\"%s\", \"%s\")", tagVal, tagVal))
			} else {
				_ = xlsx.SetCellValue(defaultSheetName, cellIndex, tagVal)
			}

			c++
		}
	}

	return xlsx, nil
}

func ExportExcelSheets(sheets []*ExcelSheet) *excelize.File {
	xlsx := excelize.NewFile()
	if len(sheets) == 0 {
		return xlsx
	}

	for _, sheet := range sheets {
		if sheet.Name == "" {
			sheet.Name = defaultSheetName
		}
		_, _ = xlsx.NewSheet(sheet.Name)

		for c, header := range sheet.Headers {
			if header.Width == 0 {
				header.Width = 20
			}
			_ = xlsx.SetColWidth(sheet.Name, columnFlags[c], columnFlags[c], header.Width)

			cellIndex := columnFlags[c] + "1"
			_ = xlsx.SetCellValue(sheet.Name, cellIndex, header.Name)
		}

		for r, data := range sheet.Datas {
			for c, val := range data {
				cellIndex := columnFlags[c] + strconv.Itoa(r+2)
				strVal, ok := val.(string)
				if ok && urlRegex.MatchString(strVal) {
					_ = xlsx.SetCellFormula(sheet.Name, cellIndex, fmt.Sprintf("=HYPERLINK(\"%s\", \"%s\")", val, val))
				} else {
					_ = xlsx.SetCellValue(sheet.Name, cellIndex, val)
				}
			}
		}
	}

	return xlsx
}
