package main

import (
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

	// xlsxImport()
	// xlsxExportSample()
	xlsxExportChart()
}
