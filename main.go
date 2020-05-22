package main

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/sirupsen/logrus"
)

func init() {
	logrus.SetLevel(logrus.DebugLevel)
	logrus.SetFormatter(&logrus.JSONFormatter{})
}

func main() {
	// csvImportLarge()
	// csvImportSmall()
	// csvExportAppend()

	xlsxImport()
}

func xlsxImport() {
	f, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		panic(err)
	}

	sheetMap := f.GetSheetMap()
	logrus.Debugf("GetSheetMap sheetMap=%+v", sheetMap)

	for mK, mV := range sheetMap {
		logrus.Debugf("(%T) %+v : (%T) %+v", mK, mK, mV, mV)
	}

	for sheetIndex := 1; sheetIndex <= len(sheetMap); sheetIndex++ {
		logrus.Debugf("sheet[%d] %s", sheetIndex, sheetMap[sheetIndex])
		strs := f.GetRows(sheetMap[sheetIndex])

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
