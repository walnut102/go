package main

import (
	"bufio"
	"fmt"
	"io"
	"io/ioutil"
	"os"
)

func main() {
	err1 := readB()
	if err1 != nil {
		panic(err1)
	}

	err2 := appendTo("a.txt", "b.txt")
	if err2 != nil {
		panic(err2)
	}

	err3 := appendTo("b.txt", "c.txt")
	if err3 != nil {
		panic(err3)
	}
}

func readB() error {
	content, err := ioutil.ReadFile("b.txt")
	if err != nil {
		fmt.Print("文件b打开失败")
		return err
	}
	fmt.Println("BRule:\n", string(content))
	return nil
}

func appendTo(file1 string, file2 string) error {
	sourceEx, err := os.Open(file1)
	if err != nil {
		fmt.Println("读取的文件打开失败")
		return err
	}
	defer sourceEx.Close()

	sourceIn, err := os.OpenFile(file2, os.O_RDWR|os.O_CREATE|os.O_APPEND, 0666)
	if err != nil {
		return err
	}
	defer sourceIn.Close()

	write := bufio.NewWriter(sourceIn)
	write.WriteString("\n")
	write.Flush()

	byteWritten, err := io.Copy(sourceIn, sourceEx)
	if err != nil {
		fmt.Println("写入的文件打开失败")
		return err
	}
	fmt.Println(byteWritten)
	return nil
}
