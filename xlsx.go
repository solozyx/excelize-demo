package main

import (
	"fmt"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/sirupsen/logrus"
)

func xlsxImport() {
	f, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		panic(err)
	}

	// Get value from cell by given worksheet name and axis.
	cell, err := f.GetCellValue("Sheet2", "B2")
	if err != nil {
		panic(err)
	}
	fmt.Println(cell)

	sheetMap := f.GetSheetMap()
	logrus.Debugf("GetSheetMap sheetMap=%+v", sheetMap)

	for mK, mV := range sheetMap {
		logrus.Debugf("(%T) %+v : (%T) %+v", mK, mK, mV, mV)
	}

	for sheetIndex := 1; sheetIndex <= len(sheetMap); sheetIndex++ {
		logrus.Debugf("sheet[%d] %s", sheetIndex, sheetMap[sheetIndex])
		strs, err := f.GetRows(sheetMap[sheetIndex])
		if err != nil {
			panic(err)
		}

		if len(strs) == 0 {
			continue
		}

		for i := 0; i < len(strs); i++ {
			logrus.Debugf("strs[%d] len=%d datas=%+v", i, len(strs[i]), strs[i])
			for j := 0; j < len(strs[i]); j++ {
				if strs[i][j] == "" {
					continue
				}
				logrus.Debugf("strs[row_%d][col_%d] type=%T data=%+v", i, j, strs[i][j], strs[i][j])
			}
		}
	}
}

func xlsxExportSample() {
	f := excelize.NewFile()

	// Create a new sheet.
	index := f.NewSheet("Sheet2")

	// Set value of a cell.
	f.SetCellValue("Sheet1", "B2", 100)
	f.SetCellValue("Sheet2", "A2", "Hello world.")

	// Set active sheet of the workbook.
	f.SetActiveSheet(index)

	// Save xlsx file by the given path.
	if err := f.SaveAs("Book2.xlsx"); err != nil {
		fmt.Println(err)
	}
}
