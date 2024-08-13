package utils

import "github.com/spf13/pflag"

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"os"
	"strings"
)

type RuleInfo struct {
	File  string
	Sheet string
	Err   error
}

func GetXlsx() RuleInfo {
	//新表格文件名
	var fileNew string
	pflag.StringVar(&fileNew, "fileNew", "b.xlsx", "储存规则的表格")
	//新单元表名
	var sheetNew string
	pflag.StringVar(&sheetNew, "sheetNew", "Sheet2", "储存规则的单元表")

	pflag.Parse()

	rule, err := os.ReadFile("rule.txt")
	if err != nil {
		fmt.Println("打开储存规则表的信息的文件失败", err)
		return RuleInfo{"", "", err}
	}
	fileOri := strings.Split(string(rule), ",")[0]
	sheetOri := strings.Split(string(rule), ",")[1]

	NewFile, err := xlsx.OpenFile(fileNew)
	if err != nil {
		fmt.Println("储存规则的打表格时出现错误:", err)
		return RuleInfo{"", "", err}
	}
	defer func() {
		err2 := NewFile.Save(fileNew)
		if err2 != nil {
			fmt.Println("储存规则的表格储存时出现错误:", err2)
		}
	}()
	NewSheet := NewFile.Sheet[sheetNew]
	OriFile, err := xlsx.OpenFile(fileOri)
	if err != nil {
		fmt.Println("原储存规则的打表格时出现错误:", err)
		return RuleInfo{"", "", err}
	}
	defer func() {
		err2 := OriFile.Save(fileOri)
		if err2 != nil {
			fmt.Println("原储存规则的表格储存时出现错误:", err2)
		}
	}()

	OriSheet := OriFile.Sheet[sheetOri]

	for i, row := range OriSheet.Rows {
		blankRow := OriSheet.AddRow()
		newRow := NewSheet.Rows[i]
		for j, cell := range row.Cells {
			blankCell := blankRow.AddCell()
			blankCell.Value = ""
			newRow.AddCell()
			newRow.Cells[j].Value = cell.Value
		}
	}

	file, err := os.Create("rule.txt")
	if err != nil {
		return RuleInfo{"", "", err}
	}
	defer file.Close()
	_, err = file.WriteString(fileNew + "," + sheetNew)
	if err != nil {
		fmt.Println(err)
		return RuleInfo{"", "", err}
	}

	return RuleInfo{fileNew, sheetNew, nil}

}
