package excel

import (
	"github.com/xuri/excelize/v2"
)

func CreateNewSheet(file *excelize.File, sheetName string, rows [][]any) (*excelize.File, error) {
	var err error
	var sheetIndex int
	if sheetIndex, err = file.NewSheet(sheetName); err != nil {
		return file, err
	} else {
		file.SetActiveSheet(sheetIndex)
		return SetCellValues(file, sheetName, rows)
	}
}
