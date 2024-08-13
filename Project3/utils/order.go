package utils

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"math"
	"strings"
)

type Order struct {
	Repeat   bool
	SourceEx string
	SourceIn string
	Detail   string
	Kind     int
	RowNum   int
	LastOri  int
}

func Clone(order *Order) (int, int) {
	fileEx, err := xlsx.OpenFile(order.SourceEx)
	if err != nil {
		fmt.Println("输出数据的表格打开失败:", err)
		return 0, 0
	}
	fileIn, err := xlsx.OpenFile(order.SourceIn)
	if err != nil {
		fmt.Println("输入数据的表格打开失败:", err)
		return 0, 0
	}

	ExAndIn := strings.Split(order.Detail, "/")
	imfSheetEx := strings.Split(ExAndIn[0], "!")
	sheetEx := fileEx.Sheet[imfSheetEx[0]]
	exOri := Convert(strings.Split(imfSheetEx[1], ":")[0])
	exDes := Convert(strings.Split(imfSheetEx[1], ":")[1])
	imfSheetIn := strings.Split(ExAndIn[1], "!")
	sheetIn := fileIn.Sheet[imfSheetIn[0]]
	inOri := Convert(strings.Split(imfSheetIn[1], ":")[0])
	inDes := Convert(strings.Split(imfSheetIn[1], ":")[1])

	//获取总行数
	var count int
	truth := true
	last := 0
	if order.Kind == 0 {
		last = exOri
	} else {
		last = order.LastOri
	}
	for i := 0; truth; i++ {
		sheetEx.AddRow()
		row := sheetEx.Rows[i]
		for j := 0; j < last; j++ {
			row.AddCell()
		}
		for j := last; j <= last; j++ {
			if row.Cells[j-1].String() == "" {
				truth = false
				count = i
				break
			}
		}
	}

	if order.Kind == 0 {
		order.RowNum = count
	}

	//获取输出数据
	data := make([][]string, 0)
	for i := count - order.RowNum; i < count; i++ {
		dataRow := make([]string, 0)
		var row *xlsx.Row
		row = sheetEx.Rows[i]
		for j := exOri; j <= exDes; j++ {
			cell := row.Cells[j-1]
			dataRow = append(dataRow, cell.String())
		}
		data = append(data, dataRow)
	}
	newRow := 0
	truth = true
	if order.Kind == 0 && order.Repeat == true {
		for i := 0; truth; i++ {
			sheetIn.AddRow()
			row := sheetIn.Rows[i]
			for j := 0; j < inOri; j++ {
				row.AddCell()
			}
			for j := inOri; j <= inOri; j++ {
				if row.Cells[j-1].String() == "" {
					truth = false
					newRow = i
					break
				}
			}
		}
	}

	line := 0
	//将data切片中的数据传入到新的表格中
	for i := count + newRow; i < count+order.RowNum+newRow; i++ {
		var row *xlsx.Row
		sheetIn.AddRow()
		row = sheetIn.Rows[i-count]
		col := 0
		for j := 0; j < inOri; j++ {
			row.AddCell()
		}
		for j := inOri; j <= inDes; j++ {
			row.AddCell()
			cell := row.Cells[j-1]
			cell.SetString(data[line][col])
			col++
		}
		line++
	}
	err = fileIn.Save(order.SourceIn)
	if err != nil {
		fmt.Println("文件储存时发生错误：", err)
		return 0, 0
	}
	return count, inOri
}

// Convert 将表格的字母转化为数字
func Convert(s string) int {
	alphabet := GetMap()
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

// GetMap 对map进行初始化,A为1
func GetMap() map[string]int {
	var alphabet = make(map[string]int)
	for i := 0; i < 26; i++ {
		var ascii = rune(65 + i)
		alphabet[string(ascii)] = i + 1
	}
	return alphabet
}

func Examine(order *Order) int {
	ExAndIn := strings.Split(order.Detail, "/")
	imfSheetEx := strings.Split(ExAndIn[0], "!")
	exOri := Convert(strings.Split(imfSheetEx[1], ":")[0])
	exDes := Convert(strings.Split(imfSheetEx[1], ":")[1])
	imfSheetIn := strings.Split(ExAndIn[1], "!")
	inOri := Convert(strings.Split(imfSheetIn[1], ":")[0])
	inDes := Convert(strings.Split(imfSheetIn[1], ":")[1])
	if exOri == 0 || exDes == 0 || inOri == 0 || inDes == 0 {
		fmt.Println("列名解析错误，请在指令中输入正确的行列")
		return 0
	}

	if (exDes - exOri) > (inDes - inOri) {
		fmt.Println("目标格的范围不匹配，请检查指令是否正确")
		return 0
	}

	return exDes - exOri
}
