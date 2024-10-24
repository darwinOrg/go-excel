package dgexcel

import (
	"fmt"
	dgcoll "github.com/darwinOrg/go-common/collection"
	"github.com/xuri/excelize/v2"
	"os"
	"reflect"
	"regexp"
	"strconv"
	"strings"
)

const defaultSheetName = "Sheet1"

var (
	columnFlags  = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	urlRegex     = regexp.MustCompile(`^((https|http|ftp|rtsp|mms)?://)\S+$`)
	nameRegex    = regexp.MustCompile(`name\((.*?)\)`)
	mappingRegex = regexp.MustCompile(`mapping\((.*?)\)`)
	widthRegex   = regexp.MustCompile(`width\((.*?)\)`)
)

type ExcelHeader struct {
	Name        string
	Width       float64
	AlignCenter bool
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
		panic(fmt.Sprintf("type %v not support", reflect.TypeOf(v).Kind()))
	}
	return resList
}

func ExportStruct2Xlsx(v any) (*excelize.File, error) {
	tagList := getStructTagList(v, "excel")
	mapTagList := struct2MapTagList(v)
	xlsx := excelize.NewFile()
	_, _ = xlsx.NewSheet(defaultSheetName)
	centerStyleId := buildCenterStyleId(xlsx)

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
		_ = xlsx.SetColWidth(defaultSheetName, columnFlags[c], columnFlags[c], float64(wt))

		cellIndex := columnFlags[c] + "1"
		_ = xlsx.SetCellValue(defaultSheetName, cellIndex, name)
		_ = xlsx.SetCellStyle(defaultSheetName, cellIndex, cellIndex, centerStyleId)
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

			cellIndex := columnFlags[c] + strconv.Itoa(r+2)
			if urlRegex.MatchString(tagVal) {
				_ = xlsx.SetCellFormula(defaultSheetName, cellIndex, fmt.Sprintf("=HYPERLINK(\"%s\", \"%s\")", tagVal, tagVal))
			} else {
				_ = xlsx.SetCellValue(defaultSheetName, cellIndex, tagVal)
			}

			c++
		}
	}

	frozenFirstRow(xlsx, defaultSheetName)

	return xlsx, nil
}

func ExportStruct2XlsxByTemplate(v any, templateFilePath string) (*excelize.File, error) {
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
	headers := rows[0]

	tagList := getStructTagList(v, "excel")
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
					cellIndex := columnFlags[c] + strconv.Itoa(r+2)
					if urlRegex.MatchString(tagVal) {
						_ = xlsx.SetCellFormula(firstSheetName, cellIndex, fmt.Sprintf("=HYPERLINK(\"%s\", \"%s\")", tagVal, tagVal))
					} else {
						_ = xlsx.SetCellValue(firstSheetName, cellIndex, tagVal)
					}

					cellStyle, err := xlsx.GetCellStyle(firstSheetName, columnFlags[c]+"2")
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
	centerStyleId := buildCenterStyleId(xlsx)

	for i, sheet := range sheets {
		if sheet.Name == "" {
			sheet.Name = defaultSheetName
		}
		if i == 0 {
			_ = xlsx.SetSheetName(defaultSheetName, sheet.Name)
		} else {
			_, _ = xlsx.NewSheet(sheet.Name)
		}

		var alignCenterColumns []int

		for c, header := range sheet.Headers {
			if header.Width == 0 {
				header.Width = 20
			}
			if header.AlignCenter {
				alignCenterColumns = append(alignCenterColumns, c)
			}

			_ = xlsx.SetColWidth(sheet.Name, columnFlags[c], columnFlags[c], header.Width)
			cellIndex := columnFlags[c] + "1"
			_ = xlsx.SetCellValue(sheet.Name, cellIndex, header.Name)
			_ = xlsx.SetCellStyle(sheet.Name, cellIndex, cellIndex, centerStyleId)
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
				if dgcoll.Contains(alignCenterColumns, c) {
					_ = xlsx.SetCellStyle(sheet.Name, cellIndex, cellIndex, centerStyleId)
				}
			}
		}

		frozenFirstRow(xlsx, sheet.Name)
	}

	return xlsx
}

func frozenFirstRow(xlsx *excelize.File, sheetName string) {
	_ = xlsx.SetPanes(sheetName, &excelize.Panes{
		Freeze:      true,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
	})
}

func buildCenterStyleId(xlsx *excelize.File) int {
	styleId, _ := xlsx.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})

	return styleId
}
