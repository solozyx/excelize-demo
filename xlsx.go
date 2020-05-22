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

func xlsxExportChart() {
	categories := map[string]string{
		"A2": "Small",
		"A3": "Normal",
		"A4": "Large",
		"B1": "Apple",
		"C1": "Orange",
		"D1": "Pear",
	}
	values := map[string]int{
		"B2": 3,
		"C2": 3,
		"D2": 3,
		"B3": 5,
		"C3": 2,
		"D3": 4,
		"B4": 6,
		"C4": 7,
		"D4": 8,
	}

	f := excelize.NewFile()
	for k, v := range categories {
		f.SetCellValue("Sheet1", k, v)
	}
	for k, v := range values {
		f.SetCellValue("Sheet1", k, v)
	}

	format := `{
		"type":"col3DClustered",
		"series":[
			{"name":"Sheet1!$A$2","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$2:$D$2"},
			{"name":"Sheet1!$A$3","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$3:$D$3"},
			{"name":"Sheet1!$A$4","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$4:$D$4"}
		],
		"title":{"name":"Fruit 3D Clustered Column Chart"}
	}`
	err := f.AddChart("Sheet1", "E1", format)
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.SaveAs("Book3.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
