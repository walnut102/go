package main

import (
	"Project2/utils"
	"fmt"
)

func main() {
	rule := utils.GetXlsx()
	if rule.Err != nil {
		return
	}
	information := utils.Read(rule.File, rule.Sheet)
	var scope [][]int
	for i, command := range information {
		if i == 0 {
			continue
		}
		num1 := utils.Examine(&utils.Order{false, command[1], "b.xlsx", command[2], 0, 0, 0})
		num2 := utils.Examine(&utils.Order{false, "b.xlsx", command[3], command[4], 1, 0, 0})
		var dataScope []int
		dataScope = append(dataScope, num1)
		dataScope = append(dataScope, num2)
		scope = append(scope, dataScope)
		for j := 1; j < i; j++ {
			if information[i-1][0] == information[j][0] {
				if scope[i-1][0] == scope[j-1][0] && scope[i-1][1] == scope[j-1][1] {
					break
				}
				fmt.Println("请检查规则格式是否正确")
				return
			}
		}
	}
	for i, command := range information {
		if i == 0 {
			continue
		}
		isRepeat := false
		for j := 0; j < i; j++ {
			if command[0] == information[j][0] {
				isRepeat = true
				break
			}
		}
		rowNum, last := utils.Clone(&utils.Order{isRepeat, command[1], "b.xlsx", command[2], 0, 0, 0})
		utils.Clone(&utils.Order{isRepeat, "b.xlsx", command[3], command[4], 1, rowNum, last})
	}

}
