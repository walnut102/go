package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
)

func main() {
	readB()

	aTob()

	bToc()
}
func readB() {
	fileB, err := xlsx.OpenFile("b.xlsx")
	if err != nil {
		fmt.Println("b文件打开时出现错误:", err)
	}
	defer func() {
		err := fileB.Save("b.xlsx")
		if err != nil {
			fmt.Println("表格b储存时出现错误:", err)
		}
	}()
	sheet := fileB.Sheets[0]
	for _, row := range sheet.Rows {
		for _, cell := range row.Cells {
			str := cell.String()
			fmt.Printf("%s\t", str)
		}
		fmt.Println()
	}
}

func aTob() {
	fileA, err := xlsx.OpenFile("a.xlsx")
	if err != nil {
		fmt.Println("表格a打开失败:", err)
		return
	}
	fileB, err := xlsx.OpenFile("b.xlsx")
	if err != nil {
		fmt.Println("表格b打开失败:", err)
		return
	}

	sheetA := fileA.Sheets[0]
	sheetB := fileB.Sheets[0]

	var count int
label:
	for i, row := range sheetB.Rows {
		for _, cell := range row.Cells {
			if cell.String() != "" {
				continue
			}
			count = i
			fmt.Println(count)
			break label
		}
	}

	for i, rowA := range sheetA.Rows {
		sheetB.AddRow()
		rowB := sheetB.Rows[i+count]
		for j, cellA := range rowA.Cells {
			rowB.AddCell()
			cellB := rowB.Cells[j]
			cellB.SetString(cellA.String())
		}
	}

	err = fileB.Save("b.xlsx")
	if err != nil {
		fmt.Println("表格b保存失败:", err)
		return
	}
}

func bToc() {
	fileB, err := xlsx.OpenFile("b.xlsx")
	if err != nil {
		fmt.Println("表格b打开失败:", err)
		return
	}
	fileC, err := xlsx.OpenFile("c.xlsx")
	if err != nil {
		fmt.Println("表格c打开失败:", err)
		return
	}

	sheetB := fileB.Sheets[0]
	sheetC := fileC.Sheets[0]

	var count int
label:
	for i, row := range sheetC.Rows {
		for _, cell := range row.Cells {
			if cell.String() != "" {
				continue
			}
			count = i
			fmt.Println(count)
			break label
		}
	}

	for i, rowB := range sheetB.Rows {
		sheetC.AddRow()
		rowC := sheetC.Rows[i+count]
		for j, cellB := range rowB.Cells {
			rowC.AddCell()
			cellC := rowC.Cells[j]
			cellC.SetString(cellB.String())
		}
	}

	err = fileC.Save("c.xlsx")
	if err != nil {
		fmt.Println("表格c保存失败:", err)
		return
	}
}
