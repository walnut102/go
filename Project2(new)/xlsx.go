package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strings"
)

var alphabet = make(map[string]int)

func main() {
	alphabet["A"] = 0
	alphabet["B"] = 1
	alphabet["C"] = 2
	alphabet["D"] = 3
	alphabet["E"] = 4
	alphabet["F"] = 5
	alphabet["G"] = 6
	alphabet["H"] = 7
	alphabet["I"] = 8
	alphabet["J"] = 9
	alphabet["K"] = 10
	alphabet["L"] = 11
	alphabet["M"] = 12
	alphabet["N"] = 13
	alphabet["O"] = 14
	alphabet["P"] = 15
	alphabet["Q"] = 16
	alphabet["R"] = 17
	alphabet["S"] = 18
	alphabet["T"] = 19
	alphabet["U"] = 20
	alphabet["V"] = 21
	alphabet["W"] = 22
	alphabet["X"] = 23
	alphabet["Y"] = 24
	alphabet["Z"] = 25

	information := readB()
	for i, command := range information {
		if i == 0 {
			continue
		}
		clone(command[1], "b.xlsx", command[2])
		clone("b.xlsx", command[3], command[4])
	}

}
func readB() [][]string {
	fileB, err := xlsx.OpenFile("b.xlsx")
	if err != nil {
		fmt.Println("b文件打开时出现错误:", err)
	}
	defer func() {
		err2 := fileB.Save("b.xlsx")
		if err2 != nil {
			fmt.Println("表格b储存时出现错误:", err2)
		}
	}()
	sheet := fileB.Sheet["Sheet2"]
	information := make([][]string, 0)
	for _, row := range sheet.Rows {
		rowData := make([]string, 0)
		for _, cell := range row.Cells {
			str := cell.String()
			rowData = append(rowData, str)
			fmt.Printf("%s\t\t", str)
		}
		fmt.Println()
		information = append(information, rowData)
	}
	return information
}

func clone(sourceEx string, sourceIn string, detail string) {
	fileEx, err := xlsx.OpenFile(sourceEx)
	if err != nil {
		fmt.Println("输出数据的表格打开失败:", err)
		return
	}

	fileIn, err := xlsx.OpenFile(sourceIn)
	if err != nil {
		fmt.Println("输入数据的表格打开失败:", err)
		return
	}

	ExAndIn := strings.Split(detail, "/")
	ex := ExAndIn[0]
	imfSheetEx := strings.Split(ex, "!")
	sheetEx := fileEx.Sheet[imfSheetEx[0]]
	//将列中的字母转换为数字
	rangeExOri := strings.Split(imfSheetEx[1], ":")[0]
	exOri := alphabet[rangeExOri]
	rangeExDes := strings.Split(imfSheetEx[1], ":")[1]
	exDes := alphabet[rangeExDes]
	in := ExAndIn[1]
	imfSheetIn := strings.Split(in, "!")
	sheetIn := fileIn.Sheet[imfSheetIn[0]]
	rangeInOri := strings.Split(imfSheetIn[1], ":")[0]
	inOri := alphabet[rangeInOri]
	rangeInDes := strings.Split(imfSheetIn[1], ":")[1]
	inDes := alphabet[rangeInDes]

	//获取输出的数据
	var count int
	data := make([][]string, 0)
	var truth = true
	for i := 0; truth; i++ {
		dataRow := make([]string, 0)
		sheetEx.AddRow()
		row := sheetEx.Rows[i]
		for j := exOri; j <= exDes; j++ {
			row.AddCell()
			if row.Cells[0].String() == "" {
				truth = false
				count = i
				break
			}
			cell := row.Cells[j]
			dataRow = append(dataRow, cell.String())
			continue

		}
		if truth {
			data = append(data, dataRow)
		}
	}

	//将data切片中的数据传入到新的表格中
	for i := 0; i < count; i++ {

		sheetIn.AddRow()
		row := sheetIn.Rows[i]
		
		col := 0
		for j := inOri; j <= inDes; j++ {
			row.AddCell()
			cell := row.Cells[j]
			cell.SetString(data[i][col])
			col++
		}
	}
	err = fileIn.Save(sourceIn)
	if err != nil {
		fmt.Println("文件储存时发生错误：", err)
		return
	}
}
