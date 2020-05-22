package main

import (
	"encoding/csv"
	"log"
	"os"

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
}

func openCsv() {
	f, err := excelize.OpenFile("Book1.csv")
	if err != nil {
		panic(err)
	}
	logrus.Debugf("f=%+v", f)
}

func openXlsx() {
	f, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		panic(err)
	}

	sheetMap := f.GetSheetMap()
	logrus.Debugf("GetSheetMap sheetMap=%+v", sheetMap)

	for mK, mV := range sheetMap {
		logrus.Debugf("(%T) %+v : (%T) %+v", mK, mK, mV, mV)
	}

	logrus.Debugf("sheet1 %s", sheetMap[1])
	strs := f.GetRows(sheetMap[1])
	for i := 0; i < len(strs); i++ {
		logrus.Debugf("strs[%d] len=%d datas=%+v", i, len(strs[i]), strs[i])
		for j := 0; j < len(strs[i]); j++ {
			if strs[i][j] == "" {
				continue
			}
			logrus.Debugf("strs[row_%d][col_%d] type=%T data=%+v", i, j, strs[i][j], strs[i][j])
		}
	}

	logrus.Debugf("sheet2 %s", sheetMap[2])
	strs = f.GetRows(sheetMap[2])
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

func exportCsv() {
	records := [][]string{
		{"first_name", "last_name", "username"},
		{"Rob", "Pike", "rob"},
		{"Ken", "Thompson", "ken"},
		{"Robert", "Griesemer", "gri"},
	}

	f, err := os.OpenFile("Book1.csv", os.O_RDWR, os.ModePerm)
	if err != nil {
		panic(err)
	}
	w := csv.NewWriter(f)

	for _, record := range records {
		if err := w.Write(record); err != nil {
			log.Fatalln("error writing record to csv:", err)
		}
	}

	// Write any buffered data to the underlying writer (standard output).
	w.Flush()

	if err := w.Error(); err != nil {
		log.Fatal(err)
	}
}

func importCsv() {
	f, err := os.OpenFile("Book1.csv", os.O_RDWR, os.ModePerm)
	if err != nil {
		panic(err)
	}
	r := csv.NewReader(f)

	strs, err := r.Read()
	if err != nil {
		panic(err)
	}

	logrus.Debugf("importCsv strs=%+v", strs)

}
