package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"log"
	"os"
)

func csvImportLarge() {
	//准备读取文件
	//fileName := "D:\\gotest\\src\\source\\test.csv"

	fileName := "Book1.csv"
	f, err := os.Open(fileName)
	if err != nil {
		log.Fatalf("can not open the file, err is %+v", err)
	}
	defer f.Close()

	r := csv.NewReader(f)
	// 针对大文件,一行一行的读取文件
	for {
		row, err := r.Read()
		if err != nil && err != io.EOF {
			log.Fatalf("can not read, err is %+v", err)
		}
		if err == io.EOF {
			break
		}
		fmt.Println(row)
	}
}

func csvImportSmall() {
	// 针对小文件,也可以一次性读取所有的文件
	// 注意,r要重新赋值,因为ReadAll是读取剩下的

	fileName := "Book1.csv"
	f, err := os.Open(fileName)
	if err != nil {
		panic(err)
	}

	r := csv.NewReader(f)
	content, err := r.ReadAll()
	if err != nil {
		log.Fatalf("can not readall, err is %+v", err)
	}

	for _, row := range content {
		fmt.Println(row)
	}
}

func csvExportAppend() {
	// 创建一个新文件
	newFileName := "Book2.csv"
	// 该打开方式,每次都会清空文件内容
	// nf, err := os.Create(newFileName)

	// 追加写
	nf, err := os.OpenFile(newFileName, os.O_RDWR|os.O_CREATE, 0666)
	if err != nil {
		log.Fatalf("can not create file, err is %+v", err)
	}
	defer nf.Close()

	_, err = nf.Seek(0, io.SeekEnd)
	if err != nil {
		log.Fatalf("can not seek file end, err is %+v", err)
	}

	w := csv.NewWriter(nf)

	// 设置属性
	w.Comma = ','
	w.UseCRLF = true

	// data
	row := []string{"1", "2", "3", "4", "5,6"}
	err = w.Write(row)
	if err != nil {
		log.Fatalf("can not write, err is %+v", err)
	}

	// 必须刷新,才能将数据写入文件
	w.Flush()

	// 一次写入多行
	var newContent [][]string
	newContent = append(newContent, []string{"1", "2", "3", "4", "5", "6"})
	newContent = append(newContent, []string{"11", "12", "13", "14", "15", "16"})
	newContent = append(newContent, []string{"21", "22", "23", "24", "25", "26"})
	err = w.WriteAll(newContent)
	if err != nil {
		log.Fatalf("can not write all, err is %+v", err)
	}
}

func csvExport() {
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
