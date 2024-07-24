package utils

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"math"
	"strings"
)

type Order struct {
	SourceEx string
	SourceIn string
	Detail   string
	Alphabet map[string]int
}

func Clone(order *Order) {
	fileEx, err := xlsx.OpenFile(order.SourceEx)
	if err != nil {
		fmt.Println("输出数据的表格打开失败:", err)
		return
	}
	fileIn, err := xlsx.OpenFile(order.SourceIn)
	if err != nil {
		fmt.Println("输入数据的表格打开失败:", err)
		return
	}

	ExAndIn := strings.Split(order.Detail, "/")
	imfSheetEx := strings.Split(ExAndIn[0], "!")
	sheetEx := fileEx.Sheet[imfSheetEx[0]]
	exOri := Convert(strings.Split(imfSheetEx[1], ":")[0], order.Alphabet)
	exDes := Convert(strings.Split(imfSheetEx[1], ":")[1], order.Alphabet)
	imfSheetIn := strings.Split(ExAndIn[1], "!")
	sheetIn := fileIn.Sheet[imfSheetIn[0]]
	inOri := Convert(strings.Split(imfSheetIn[1], ":")[0], order.Alphabet)
	inDes := Convert(strings.Split(imfSheetIn[1], ":")[1], order.Alphabet)

	if exOri == 0 || exDes == 0 || inOri == 0 || inDes == 0 {
		fmt.Println("列名解析错误，请在指令中输入正确的行列")
		return
	}

	if (exDes - exOri) > (inDes - inOri) {
		fmt.Println("目标格的范围不匹配，请检查指令是否正确")
		return
	}
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
	err = fileIn.Save(order.SourceIn)
	if err != nil {
		fmt.Println("文件储存时发生错误：", err)
		return
	}
}

// Convert 将表格的字母转化为数字
func Convert(s string, alphabet map[string]int) int {
	n := len(s)
	sum := 0.0
	for i := 0; i < n; i++ {
		if alphabet[string(rune(s[i]))] == 0 {
			return 0
		}
		sum += math.Pow(26.0, float64(n-1-i)) * float64(alphabet[string(rune(s[i]))])
	}
	return int(sum)
}

// SetMap 对map进行初始化
func SetMap(alphabet *map[string]int) {
	for i := 0; i < 26; i++ {
		var ascii = rune(65 + i)
		(*alphabet)[string(ascii)] = i + 1
	}
}
