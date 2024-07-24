package excel

import (
	"github.com/xuri/excelize/v2"
	"strconv"
)

func SetCellValues(file *excelize.File, sheetName string, rows [][]any) (*excelize.File, error) {
	var err error
	for rowKey, row := range rows {
		for columnKey, cellValue := range row {
			if err = file.SetCellValue(sheetName, GetColumnName(columnKey+1)+strconv.Itoa(rowKey+1), cellValue); err != nil {
				return file, err
			}
		}
	}
	return file, nil
}
