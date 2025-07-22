package dgexcel

import (
	"errors"
	"fmt"
	dgctx "github.com/darwinOrg/go-common/context"
	dglogger "github.com/darwinOrg/go-logger"
	"github.com/xuri/excelize/v2"
	"os"
	"reflect"
	"strconv"
	"strings"
)

func ExportStruct2Xlsx(v any) (*excelize.File, error) {
	tagList := getStructTagList(v, excelTag)
	mapTagList := struct2MapTagList(v)
	xlsx := excelize.NewFile()
	_, _ = xlsx.NewSheet(DefaultSheetName)
	centerStyleId := BuildCenterStyleId(xlsx)

	for c, tagVal := range tagList {
		name, _ := stringMatchExport(tagVal, nameRegex)

		width, _ := stringMatchExport(tagVal, widthRegex)
		if width == "" {
			width = "20"
		}
		wt, _ := strconv.Atoi(width)
		if wt == 0 {
			wt = 20
		}
		_ = xlsx.SetColWidth(DefaultSheetName, ColumnIndexToName(c), ColumnIndexToName(c), float64(wt))

		cellIndex := ColumnIndexToName(c) + "1"
		_ = xlsx.SetCellValue(DefaultSheetName, cellIndex, name)
		_ = xlsx.SetCellStyle(DefaultSheetName, cellIndex, cellIndex, centerStyleId)
	}

	for r, mapTagVal := range mapTagList {
		c := 0
		for i, tagVal := range mapTagVal {
			tagKey := tagList[i]
			mapping, _ := stringMatchExport(tagKey, mappingRegex)
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

			cellIndex := ColumnIndexToName(c) + strconv.Itoa(r+2)
			if urlRegex.MatchString(tagVal) {
				_ = xlsx.SetCellFormula(DefaultSheetName, cellIndex, fmt.Sprintf("=HYPERLINK(\"%s\", \"%s\")", tagVal, tagVal))
			} else {
				_ = xlsx.SetCellValue(DefaultSheetName, cellIndex, tagVal)
			}

			c++
		}
	}

	FrozenFirstRow(xlsx, DefaultSheetName)

	return xlsx, nil
}

func ExportStruct2XlsxByTemplate(v any, templateFilePath string, headerRow int) (*excelize.File, error) {
	file, err := os.Open(templateFilePath)
	if err != nil {
		return nil, err
	}
	defer func(f *os.File) {
		_ = f.Close()
	}(file)
	xlsx, err := excelize.OpenReader(file)
	if err != nil {
		return nil, err
	}

	firstSheetName := xlsx.GetSheetList()[0]
	rows, err := xlsx.GetRows(firstSheetName)
	if err != nil {
		return nil, err
	}
	if len(rows) < headerRow+1 {
		return nil, errors.New("invalid header row")
	}
	headers := rows[headerRow]

	tagList := getStructTagList(v, excelTag)
	mapTagList := struct2MapTagList(v)

	for r, mapTagVal := range mapTagList {
		for i, tagVal := range mapTagVal {
			tagKey := tagList[i]
			mapping, _ := stringMatchExport(tagKey, mappingRegex)
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

			name, _ := stringMatchExport(tagKey, nameRegex)
			if name == "" {
				continue
			}

			for c, header := range headers {
				if header == name {
					cellIndex := ColumnIndexToName(c) + strconv.Itoa(r+2)
					if urlRegex.MatchString(tagVal) {
						_ = xlsx.SetCellFormula(firstSheetName, cellIndex, fmt.Sprintf("=HYPERLINK(\"%s\", \"%s\")", tagVal, tagVal))
					} else {
						_ = xlsx.SetCellValue(firstSheetName, cellIndex, tagVal)
					}

					cellStyle, err := xlsx.GetCellStyle(firstSheetName, ColumnIndexToName(c)+"2")
					if err == nil {
						_ = xlsx.SetCellStyle(firstSheetName, cellIndex, cellIndex, cellStyle)
					}

					break
				}
			}
		}
	}

	return xlsx, nil
}

func ExportExcelSheets(sheets []*ExcelSheet) *excelize.File {
	xlsx := excelize.NewFile()
	if len(sheets) == 0 {
		return xlsx
	}

	for i, sheet := range sheets {
		if sheet.Name == "" {
			sheet.Name = fmt.Sprintf("Sheet%d", xlsx.SheetCount+1)
		}
		if i == 0 {
			_ = xlsx.SetSheetName(DefaultSheetName, sheet.Name)
		} else {
			_, _ = xlsx.NewSheet(sheet.Name)
		}
	}

	FillExcelSheets(xlsx, sheets)
	return xlsx
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
		dglogger.Errorf(dgctx.SimpleDgContext(), "type %v not support", reflect.TypeOf(v).Kind())
		return []string{}
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

func getTagValMap(v any) []string {
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
		rv := reflect.ValueOf(v)
		if isPtr {
			rv = rv.Elem()
		}
		val := rv.FieldByName(structField.Name)
		resMap = append(resMap, fmt.Sprintf("%v", val.Interface()))
	}

	return resMap
}

func struct2MapTagList(v any) [][]string {
	var resList [][]string
	switch reflect.TypeOf(v).Kind() {
	case reflect.Slice, reflect.Array:
		values := reflect.ValueOf(v)
		for i := 0; i < values.Len(); i++ {
			resList = append(resList, getTagValMap(values.Index(i).Interface()))
		}
		break
	case reflect.Struct:
		val := reflect.ValueOf(v).Interface()
		resList = append(resList, getTagValMap(val))
		break
	default:
		dglogger.Errorf(dgctx.SimpleDgContext(), "type %v not support", reflect.TypeOf(v).Kind())
	}
	return resList
}
