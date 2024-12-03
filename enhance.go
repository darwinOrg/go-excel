package dgexcel

import (
	dgerr "github.com/darwinOrg/go-common/enums/error"
	"github.com/xuri/excelize/v2"
	"slices"
	"strings"
)

func RemoveDuplicateRowsByColumn(xlsx *excelize.File, sheetName string, startRowIndex int, columnIndex int) error {
	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return err
	}
	if startRowIndex >= len(rows) {
		return dgerr.ARGUMENT_NOT_VALID
	}

	rows = rows[startRowIndex:]
	mp := make(map[string]bool)
	var duplicatedRowIndexes []int

	for i, row := range rows {
		cell := strings.TrimSpace(row[columnIndex])
		if cell == "" {
			continue
		}

		if mp[cell] {
			duplicatedRowIndexes = append(duplicatedRowIndexes, i)
		} else {
			mp[cell] = true
		}
	}

	if len(duplicatedRowIndexes) > 0 {
		slices.Reverse(duplicatedRowIndexes)
		for _, rowIndex := range duplicatedRowIndexes {
			err = xlsx.RemoveRow(sheetName, rowIndex)
			if err != nil {
				return err
			}
		}
	}

	return nil
}
