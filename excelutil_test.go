package main

import (
	"fmt"
	"testing"
)

type User struct {
	Id    int    `excel:"编号"`
	Name  string `excel:"学生名字"`
	Age   int    `excel:"年龄"`
	Class string `excel:"班级"`
}

func TestGenExcel(t *testing.T) {
	var users []User
	user1 := User{
		Id:    1,
		Name:  "兔瓜",
		Age:   11,
		Class: "A1",
	}
	user2 := User{
		Id:    2,
		Name:  "兔柠",
		Age:   13,
		Class: "A1",
	}
	users = append(users, user1)
	users = append(users, user2)
	file, err := GenExcel(users, 2)
	if err != nil {
		fmt.Printf("%v\n", err)
	}
	file.SetCellValue("sheet1", "A1", "兔学院学生表")
	err = file.SaveAs("user.xlsx")
	if err != nil {
		fmt.Printf("%v\n", err)
	}
}
