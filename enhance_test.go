package dgexcel

import (
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestMergeDuplicateCellsByColumn(t *testing.T) {
	xlsx, err := excelize.OpenFile("test.xlsx")
	if err != nil {
		panic(err)
	}
	MergeDuplicateCellsByColumn(xlsx, DefaultSheetName, 0)
	_ = xlsx.SaveAs("test_merged.xlsx")
}
