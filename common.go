package dgexcel

import (
	"fmt"
	dgcoll "github.com/darwinOrg/go-common/collection"
	"github.com/darwinOrg/go-common/utils"
	"github.com/xuri/excelize/v2"
	"regexp"
	"strconv"
)

const DefaultSheetName = "Sheet1"

const (
	excelTag   = "excel"
	nameTag    = "name"
	uniqueTag  = "unique"
	dateTag    = "date"
	mappingTag = "mapping"
)

var (
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

func ColumnIndexToName(index int) string {
	name, err := excelize.ColumnNumberToName(index + 1)
	if err != nil {
		return "A"
	}
	return name
}

func WriteCell(xlsx *excelize.File, sheetName string, row, col, styleId int, data any) {
	colStr := ColumnIndexToName(col)
	rowStr := strconv.Itoa(row + 1)
	cell := colStr + rowStr

	_ = xlsx.SetCellValue(sheetName, cell, data)

	if styleId > 0 {
		_ = xlsx.SetCellStyle(sheetName, cell, cell, styleId)
	}
}

func WriteRowStruct(xlsx *excelize.File, sheetName string, row, fromCol, styleId int, obj any) {
	datas := utils.ReflectAllFieldValues(obj)
	WriteRowDatas(xlsx, sheetName, row, fromCol, styleId, datas...)
}

func WriteRowDatas(xlsx *excelize.File, sheetName string, row, fromCol, styleId int, datas ...any) {
	rowStr := strconv.Itoa(row + 1)

	for i, data := range datas {
		_ = xlsx.SetCellValue(sheetName, ColumnIndexToName(fromCol+i)+rowStr, data)
	}

	if styleId > 0 {
		_ = xlsx.SetCellStyle(sheetName, ColumnIndexToName(fromCol)+rowStr, ColumnIndexToName(fromCol+len(datas)-1)+rowStr, styleId)
	}
}

func WriteColumnStruct(xlsx *excelize.File, sheetName string, col, fromRow, styleId int, obj any) {
	datas := utils.ReflectAllFieldValues(obj)
	WriteColumnDatas(xlsx, sheetName, col, fromRow, styleId, datas...)
}

func WriteColumnDatas(xlsx *excelize.File, sheetName string, col, fromRow, styleId int, datas ...any) {
	colStr := ColumnIndexToName(col)

	for i, data := range datas {
		_ = xlsx.SetCellValue(sheetName, colStr+strconv.Itoa(fromRow+i), data)
	}

	if styleId > 0 {
		_ = xlsx.SetCellStyle(sheetName, colStr+strconv.Itoa(fromRow), colStr+strconv.Itoa(fromRow+len(datas)-1), styleId)
	}
}

func WriteAndMergeCell(xlsx *excelize.File, sheetName string, topLeftCell, bottomRightCell string, styleId int, data any) {
	_ = xlsx.SetCellValue(sheetName, topLeftCell, data)

	if styleId > 0 {
		_ = xlsx.SetCellStyle(sheetName, topLeftCell, bottomRightCell, styleId)
	}

	_ = xlsx.MergeCell(sheetName, topLeftCell, bottomRightCell)
}

func FillExcelSheets(xlsx *excelize.File, sheets []*ExcelSheet) {
	centerStyleId := BuildCenterStyleId(xlsx)

	for _, sheet := range sheets {
		var alignCenterColumns []int

		for c, header := range sheet.Headers {
			if header.Width == 0 {
				header.Width = 20
			}
			if header.AlignCenter {
				alignCenterColumns = append(alignCenterColumns, c)
			}

			_ = xlsx.SetColWidth(sheet.Name, ColumnIndexToName(c), ColumnIndexToName(c), header.Width)
			cellIndex := ColumnIndexToName(c) + "1"
			_ = xlsx.SetCellValue(sheet.Name, cellIndex, header.Name)
			_ = xlsx.SetCellStyle(sheet.Name, cellIndex, cellIndex, centerStyleId)
		}

		for r, data := range sheet.Datas {
			for c, val := range data {
				cellIndex := ColumnIndexToName(c) + strconv.Itoa(r+2)
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

		FrozenFirstRow(xlsx, sheet.Name)
	}
}

func AppendExcelSheets(xlsx *excelize.File, sheets []*ExcelSheet) {
	if len(sheets) == 0 {
		return
	}

	for _, sheet := range sheets {
		if sheet.Name == "" {
			sheet.Name = fmt.Sprintf("Sheet%d", xlsx.SheetCount+1)
		}
		_, _ = xlsx.NewSheet(sheet.Name)
	}

	FillExcelSheets(xlsx, sheets)
}

func BuildCenterStyleId(xlsx *excelize.File) int {
	styleId, _ := xlsx.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})

	return styleId
}

func FrozenFirstRow(xlsx *excelize.File, sheetName string) {
	FrozenRow(xlsx, sheetName, 1)
}

func FrozenRow(xlsx *excelize.File, sheetName string, row int) {
	_ = xlsx.SetPanes(sheetName, &excelize.Panes{
		Freeze:      true,
		XSplit:      0,
		YSplit:      row,
		TopLeftCell: fmt.Sprintf("A%d", row+1),
		ActivePane:  "bottomLeft",
	})
}
