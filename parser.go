package dgexcel

import (
	"encoding/json"
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path"
	"reflect"
	"regexp"
	"strconv"
	"strings"
	"time"
	"unsafe"
)

var AllowMaxRow = 10000

type parser struct {
	file         *excelize.File
	fieldMapping map[string]map[string]string
	sheetName    string
	body         any
	val          reflect.Value
	uniqueMap    map[int][]string
}

func newParser(body any) (*parser, error) {
	p := new(parser)
	p.val = reflect.ValueOf(body)
	if p.val.Kind() != reflect.Ptr {
		return nil, errors.New("body must be pointer struct")
	}
	p.body = body

	p.fieldMapping = make(map[string]map[string]string)
	//生成结构体与excel头映射关系
	p.generateMapping(p.val, "")
	return p, nil
}

func (p *parser) generateMapping(val reflect.Value, baseField string) {
	switch val.Kind() {
	case reflect.Struct:
	case reflect.Ptr:
		//当结构体指针或字段指针为空，则创建一个指针指向
		if val.IsNil() {
			newValue := reflect.New(val.Type().Elem())
			val = reflect.NewAt(val.Type().Elem(), unsafe.Pointer(newValue.Pointer()))
		}
		val = val.Elem()
		p.generateMapping(val, baseField)
		return
	default:
		return
	}
	typ := val.Type()
	for i := 0; i < val.NumField(); i++ {
		fieldName := typ.Field(i).Name
		if baseField != "" {
			fieldName = fmt.Sprintf("%s.%s", baseField, fieldName)
		}
		excel, ok := typ.Field(i).Tag.Lookup(excelTag)
		if !ok {
			//生成嵌套结构体的映射关系
			fieldVal := val.Field(i)
			p.generateMapping(fieldVal, fieldName)
			continue
		}
		m := map[string]string{nameTag: fieldName}
		m[mappingTag], _ = stringMatchExport(excel, regexp.MustCompile(`mapping\((.*?)\)`))
		m[uniqueTag], _ = stringMatchExport(excel, regexp.MustCompile(`unique\((.*?)\)`))
		m[dateTag], _ = stringMatchExport(excel, regexp.MustCompile(`date\((.*?)\)`))
		mappingName, _ := stringMatchExport(excel, regexp.MustCompile(`name\((.*?)\)`))
		p.fieldMapping[strings.TrimSpace(mappingName)] = m
	}
}

func stringMatchExport(str string, reg *regexp.Regexp) (res string, err error) {
	defer func() {
		if panicInfo := recover(); panicInfo != nil {
			err = errors.New("not match regexp")
		}
	}()
	return reg.FindStringSubmatch(str)[1], nil
}

func (p *parser) ParseContent(file *os.File, mappingHeaderRow int, dataStartRow int) (*Result, error) {
	if mappingHeaderRow-1 < 0 {
		return nil, errors.New("no excel mapping header position is specified")
	}
	if mappingHeaderRow >= dataStartRow {
		return nil, errors.New("mapping header row position cannot be greater than or equal to the beginning of the data row")
	}
	if err := p.readExcel(file); err != nil {
		return nil, err
	}
	p.uniqueMap = make(map[int][]string)
	p.sheetName = p.file.GetSheetName(0)
	rows, err := p.file.GetRows(p.sheetName)
	if err != nil {
		return nil, err
	}
	if len(rows) < dataStartRow {
		return nil, errors.New("excel file valid data behavior is empty")
	}
	//excel数据行数限制
	if len(rows)-(dataStartRow-1) > AllowMaxRow {
		return nil, errors.New("data overrun")
	}

	res := new(Result)
	res.mappingResults = make([]any, 0)
	if err := p.rows(rows, mappingHeaderRow, dataStartRow, res); err != nil {
		return nil, err
	}
	return res, nil
}

func (p *parser) readExcel(file *os.File) (err error) {
	var allowExtMap = map[string]bool{
		".xlsx": true,
	}
	ext := path.Ext(file.Name())
	//判断文件后缀
	if _, ok := allowExtMap[ext]; !ok {
		return fmt.Errorf("file request format error，support XLSX")
	}
	p.file, err = excelize.OpenReader(file)
	return err
}

func (p *parser) rows(rows [][]string, mappingHeaderRow, dataStartRow int, res *Result) error {
	for rowIndex := dataStartRow - 1; rowIndex < len(rows); rowIndex++ {
		res.rowIndex = rowIndex
		errList := make([]string, 0)
		newBodyVal := reflect.New(p.val.Type().Elem())
		newBodyVal.Elem().Set(p.val.Elem())
		for colIndex, col := range rows[rowIndex] {
			if colIndex >= len(rows[mappingHeaderRow-1]) {
				continue
			}
			mappingHeader := rows[mappingHeaderRow-1][colIndex]
			//去除列的前后空格
			colVal := strings.TrimSpace(col)
			mappingField, ok := p.fieldMapping[strings.TrimSpace(mappingHeader)]
			if !ok {
				continue
			}
			// 列唯一性校验
			errList = append(errList, p.uniqueFormat(rows, mappingHeader, &colVal, rowIndex, colIndex, mappingField)...)
			//格式化时间
			errList = append(errList, p.dateFormat(mappingHeader, &colVal, mappingField)...)
			//值映射转换
			mappingErrList := p.mappingFormat(mappingHeader, &colVal, mappingField)
			errList = append(errList, mappingErrList...)
			if len(mappingErrList) != 0 {
				continue
			}
			//参数赋值
			errs, err := p.parseValue(newBodyVal, mappingField[nameTag], mappingHeader, colVal)
			if err != nil {
				return err
			}
			errList = append(errList, errs...)
		}
		if len(errList) != 0 {
			if res.errors == nil {
				res.errors = map[int][]string{rowIndex + 1: errList}
			} else {
				res.errors[rowIndex+1] = errList
			}
		}
		p.body = newBodyVal.Interface()
		if _, ok := res.HasError(); ok {
			continue
		}
		res.mappingResults = append(res.mappingResults, p.body)
	}
	return nil
}

func (p *parser) uniqueFormat(rows [][]string, mappingHeader string, col *string, rowIndex, colIndex int, mappingField map[string]string) []string {
	errList := make([]string, 0)
	format, ok := mappingField[uniqueTag]
	if !ok || format != "true" {
		return errList
	}
	_, ok = p.uniqueMap[colIndex]
	if !ok {
		p.uniqueMap[colIndex] = make([]string, 0)
		for index := 0; index < len(rows); index++ {
			if len(rows[index]) <= colIndex {
				p.uniqueMap[colIndex] = append(p.uniqueMap[colIndex], rows[index][0])
				continue
			}
			p.uniqueMap[colIndex] = append(p.uniqueMap[colIndex], rows[index][colIndex])
		}
	}
	cols := p.uniqueMap[colIndex]
	for i, val := range cols {
		if i != rowIndex && val != "" && val == *col {
			errList = append(errList, fmt.Sprintf("%s[%s]不可重复", mappingHeader, *col))
			break
		}
	}
	return errList
}

func (p *parser) dateFormat(mappingHeader string, col *string, mappingField map[string]string) []string {
	errList := make([]string, 0)
	format, ok := mappingField[dateTag]
	if !ok || format == "" {
		return errList
	}
	formats := strings.SplitN(format, ",", 2)
	if *col == "" || len(formats) != 2 {
		return errList
	}
	location, err := time.ParseInLocation(formats[0], *col, time.Local)
	if err != nil {
		errList = append(errList, fmt.Sprintf("%s单元格格式错误", mappingHeader))
		return errList
	}
	*col = location.Format(formats[1])
	return errList
}

func (p *parser) mappingFormat(mappingHeader string, col *string, mappingField map[string]string) []string {
	errList := make([]string, 0)
	format, ok := mappingField[mappingTag]
	if !ok || format == "" {
		return errList
	}
	mappingValues := make(map[string]string)
	formatStr := strings.Split(format, ",")
	for _, format := range formatStr {
		n := strings.SplitN(format, ":", 2)
		if len(n) != 2 {
			continue
		}
		mappingValues[n[0]] = n[1]
	}
	val, ok := mappingValues[*col]
	if ok {
		*col = val
		return errList
	}
	errList = append(errList, fmt.Sprintf("%s单元格存在非法输入", mappingHeader))
	return errList
}

func (p *parser) parseValue(val reflect.Value, fieldAddr, mappingHeader, col string) ([]string, error) {
	errList := make([]string, 0)
	fields := strings.Split(fieldAddr, ".")
	if len(fields) == 0 {
		return errList, nil
	}
	for _, field := range fields {
		if val.Kind() == reflect.Ptr {
			val = val.Elem()
		}
		val = val.FieldByName(field)
		errs, err := p.parse(val, col, mappingHeader)
		if err != nil {
			return errList, err
		}
		errList = append(errList, errs...)
	}
	return errList, nil
}

func (p *parser) parse(val reflect.Value, col, mappingHeader string) ([]string, error) {
	errList := make([]string, 0)
	var err error
	switch val.Kind() {
	case reflect.String:
		val.SetString(col)
	case reflect.Bool:
		parseBool, err := strconv.ParseBool(col)
		if err != nil {
			errList = append(errList, fmt.Sprintf("%s单元格非法输入,参数非bool类型值", mappingHeader))
		}
		val.SetBool(parseBool)
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		var value int64
		if col != "" {
			value, err = strconv.ParseInt(col, 10, 64)
			if err != nil {
				errList = append(errList, fmt.Sprintf("%s单元格非法输入,参数非整形数值", mappingHeader))
			}
		}
		val.SetInt(value)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		var value uint64
		if col != "" {
			value, err = strconv.ParseUint(col, 10, 64)
			if err != nil {
				errList = append(errList, fmt.Sprintf("%s单元格非法输入,参数非整形数值", mappingHeader))
			}
		}
		val.SetUint(value)
	case reflect.Float32, reflect.Float64:
		var value float64
		if col != "" {
			value, err = strconv.ParseFloat(col, 64)
			if err != nil {
				errList = append(errList, fmt.Sprintf("%s单元格非法输入,参数非浮点型数值", mappingHeader))
			}
		}
		val.SetFloat(value)
	case reflect.Struct:
		return errList, nil
	case reflect.Ptr:
		//初始化指针
		value := reflect.New(val.Type().Elem())
		val.Set(value)
		var errs []string
		errs, err = p.parse(val.Elem(), col, mappingHeader)
		if err != nil {
			break
		}
		errList = append(errList, errs...)
	default:
		return errList, fmt.Errorf("excel column[%s] parseValue unsupported type[%v] mappings", mappingHeader, val.Kind().String())
	}
	return errList, nil
}

type Result struct {
	errors         map[int][]string
	mappingResults []any
	rowIndex       int
}

func (r *Result) HasError() (map[int][]string, bool) {
	return r.errors, len(r.errors) != 0
}

func (r *Result) List() []any {
	return r.mappingResults
}

func (r *Result) Format(ts any) error {
	marshal, err := json.Marshal(r.mappingResults)
	if err != nil {
		return err
	}
	return json.Unmarshal(marshal, &ts)
}

func (r *Result) FormatBaseTargetBuilder(buildFn func() any) ([]any, error) {
	var ret []any

	for _, elem := range r.mappingResults {
		marshal, err := json.Marshal(elem)
		if err != nil {
			return nil, err
		}
		v := buildFn()

		err = json.Unmarshal(marshal, v)
		if err != nil {
			return nil, err
		}
		ret = append(ret, v)
	}

	return ret, nil
}
