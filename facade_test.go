package dgexcel

import (
	"encoding/json"
	"fmt"
	dgctx "github.com/darwinOrg/go-common/context"
	dglogger "github.com/darwinOrg/go-logger"
	"github.com/xuri/excelize/v2"
	"testing"
)

type User struct {
	Name        string `excel:"name(姓名);unique(true);width(20)"`
	Status      int    `excel:"name(状态);mapping(无效:0,有效:1);width(40)"`
	CreatedDate string `excel:"name(创建日期);date(01-02-06,2006-01-02);width(80)"`
}

func TestSimpleBindExcel2Struct(t *testing.T) {
	ctx := &dgctx.DgContext{TraceId: "123"}
	users, err := SimpleBindExcel2Struct[User](ctx, "./users.xlsx")
	if err != nil {
		dglogger.Errorf(ctx, "bind excel to struct error: \n%v", err)
		return
	}
	usersBytes, _ := json.Marshal(users)
	dglogger.Infof(ctx, "%s", string(usersBytes))
}

func TestBindExcelOfPrototype(t *testing.T) {
	ctx := &dgctx.DgContext{TraceId: "123"}
	users, _ := BindExcelUsingTargetBuilder(ctx, "./users.xlsx", 1, 2, func() any {
		return &User{}
	})
	usersBytes, _ := json.Marshal(users)
	dglogger.Infof(ctx, "%s", string(usersBytes))
}

func TestSimpleExportStruct2XlsxFile(t *testing.T) {
	ctx := &dgctx.DgContext{TraceId: "123"}
	users, _ := SimpleBindExcel2Struct[User](ctx, "./users.xlsx")
	err := ExportStruct2XlsxFile(ctx, users, "./exported_users.xlsx")
	if err != nil {
		return
	}
}

func TestExportStruct2XlsxFileAndInsertRows(t *testing.T) {
	ctx := &dgctx.DgContext{TraceId: "123"}
	users, _ := SimpleBindExcel2Struct[User](ctx, "./users.xlsx")

	xlsx, err := ExportStruct2Xlsx(users)
	if err != nil {
		panic(err)
	}
	xlsx.InsertRows("Sheet1", 1, 1)
	xlsx.SetCellStr("Sheet1", "A1", "头部")
	//xlsx.SetCellFormula("Sheet1", "A1", "=HYPERLINK(\"https://www.baidu.com\", \"https://www.baidu.com\")")

	xlsx.SetPanes("Sheet1", &excelize.Panes{
		Freeze:      true,
		XSplit:      0,
		YSplit:      2,
		TopLeftCell: "A3",
		ActivePane:  "bottomLeft",
	})

	err = xlsx.SaveAs("./exported_users.xlsx")
	if err != nil {
		panic(err)
	}
}

func TestExportStruct2XlsxByTemplate(t *testing.T) {
	ctx := &dgctx.DgContext{TraceId: "123"}
	users, _ := SimpleBindExcel2Struct[User](ctx, "./users.xlsx")

	xlsx, err := ExportStruct2XlsxByTemplate(users, "./users_template.xlsx", 0)
	if err != nil {
		panic(err)
	}

	err = xlsx.SaveAs("./exported_users.xlsx")
	if err != nil {
		panic(err)
	}
}

func TestAddChart(t *testing.T) {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	f.NewSheet("Sheet2")

	for idx, row := range [][]interface{}{
		{"整体数据"},
		{"已提交面试", "待反馈", "待定", "未通过", "已通过"},
		{90, 20, 5, 40, 25},
	} {
		cell, err := excelize.CoordinatesToCellName(1, idx+1)
		if err != nil {
			fmt.Println(err)
			return
		}
		f.SetSheetRow("Sheet2", cell, &row)
	}
	varyColors := false
	if err := f.AddChart("Sheet1", "F1", &excelize.Chart{
		Type: excelize.Col3DClustered,
		Series: []excelize.ChartSeries{
			{
				Name:       "Sheet2!$A$1",
				Categories: "Sheet2!$A$2:$E$2",
				Values:     "Sheet2!$A$3:$E$3",
			},
		},
		Dimension: excelize.ChartDimension{
			Width:  800,
			Height: 400,
		},
		VaryColors: &varyColors,
		Title: []excelize.RichTextRun{
			{
				Text: "全部职位（2024-03-14 ~ 2024-04-14）",
			},
		},
	}); err != nil {
		fmt.Println(err)
		return
	}
	// Save spreadsheet by the given path.
	if err := f.SaveAs("chart.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func TestExportExcelSheets(t *testing.T) {
	f := ExportExcelSheets([]*ExcelSheet{
		{
			Name: "用户1",
			Headers: []*ExcelHeader{
				{Name: "姓名", Width: 20},
				{Name: "状态", Width: 20, AlignCenter: true},
				{Name: "创建日期", Width: 40},
			},
			Datas: [][]any{
				{"张三", 1, "2024-03-11"},
				{"李四", 0, "2024-04-12"},
			},
		},
		{
			Name: "用户2",
			Headers: []*ExcelHeader{
				{Name: "姓名", Width: 20},
				{Name: "状态", Width: 20},
				{Name: "创建日期", Width: 40},
			},
			Datas: [][]any{
				{"王五", 1, "2024-03-13"},
				{"赵六", 0, "2024-04-14"},
			},
		},
	})

	if err := f.SaveAs("users.xlsx"); err != nil {
		fmt.Println(err)
	}
}
