package dgexcel

import (
	"github.com/xuri/excelize/v2"
	"testing"
)

func TestWriteRowDatas(t *testing.T) {
	xlsx := excelize.NewFile()
	_, _ = xlsx.NewSheet(DefaultSheetName)
	centerStyleId := BuildCenterStyleId(xlsx)
	WriteRowDatas(xlsx, DefaultSheetName, 3, 4, centerStyleId, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
	_ = xlsx.SaveAs("test.xlsx")
}

func TestWriteColumnDatas(t *testing.T) {
	xlsx := excelize.NewFile()
	_, _ = xlsx.NewSheet(DefaultSheetName)
	centerStyleId := BuildCenterStyleId(xlsx)
	WriteColumnDatas(xlsx, DefaultSheetName, 3, 4, centerStyleId, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
	_ = xlsx.SaveAs("test.xlsx")
}
