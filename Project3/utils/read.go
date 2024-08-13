package utils

import (
	"fmt"
	"github.com/tealeg/xlsx"
)

func Read(fileName string, sheetName string) [][]string {
	file, err := xlsx.OpenFile(fileName)
	if err != nil {
		fmt.Println("规则表打开时出现错误:", err)
	}
	defer func() {
		err2 := file.Save(fileName)
		if err2 != nil {
			fmt.Println("规则表储存时出现错误:", err2)
		}
	}()
	sheet := file.Sheet[sheetName]
	information := make([][]string, 0)
	for _, row := range sheet.Rows {
		rowData := make([]string, 0)
		for _, cell := range row.Cells {
			str := cell.String()
			rowData = append(rowData, str)
			fmt.Printf("%s\t\t", str)
		}
		if rowData[0] != "" {
			fmt.Println()
		}
		information = append(information, rowData)
	}
	return information
}
