package dgexcel

import (
	"fmt"
	dgerr "github.com/darwinOrg/go-common/enums/error"
	"github.com/xuri/excelize/v2"
	"strings"
)

func RemoveDuplicateRowsByColumn(xlsx *excelize.File, sheetName string, startRowIndex int, columnIndex int, reserveValues ...string) (*excelize.File, error) {
	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return nil, err
	}
	if startRowIndex >= len(rows) {
		return nil, dgerr.ARGUMENT_NOT_VALID
	}

	newRows := rows[:startRowIndex]
	rows = rows[startRowIndex:]
	mp := make(map[string]bool)
	var duplicatedRowIndexes []int

	for i, row := range rows {
		cell := strings.TrimSpace(row[columnIndex])
		if len(reserveValues) > 0 {
			for _, reserveValue := range reserveValues {
				if cell == reserveValue {
					newRows = append(newRows, row)
					continue
				}
			}
		}

		if mp[cell] {
			duplicatedRowIndexes = append(duplicatedRowIndexes, startRowIndex+i)
		} else {
			mp[cell] = true
			newRows = append(newRows, row)
		}
	}

	if len(duplicatedRowIndexes) == 0 {
		return xlsx, nil
	}

	nf := excelize.NewFile()
	_, err = nf.NewSheet(sheetName)
	if err != nil {
		return nil, err
	}

	for rowIndex, row := range newRows {
		for colIndex, cellValue := range row {
			_ = nf.SetCellValue(sheetName, fmt.Sprintf("%s%d", ColumnIndexToName(colIndex), rowIndex+1), cellValue)
		}
	}

	return nf, nil
}
