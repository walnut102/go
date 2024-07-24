package main

import "Project2/utils"

func main() {
	var alphabet = make(map[string]int)
	utils.SetMap(&alphabet)

	information := utils.Read()
	for i, command := range information {
		if i == 0 {
			continue
		}
		utils.Clone(&utils.Order{command[1], "b.xlsx", command[2], alphabet})
		utils.Clone(&utils.Order{"b.xlsx", command[3], command[4], alphabet})
	}

}
