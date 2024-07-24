package utils

import (
	"fmt"
	"github.com/tealeg/xlsx"
)

func Read() [][]string {
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
