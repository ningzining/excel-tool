package main

import (
	"fmt"
	"testing"
)

type User struct {
	Id    int    `excel:"编号" json:"id"`
	Name  string `excel:"学生名字" json:"name"`
	Age   int    `excel:"年龄" json:"age"`
	Class string `excel:"班级" json:"class"`
}

func TestGenExcel(t *testing.T) {
	var users []*User
	user1 := User{
		Id:    1,
		Name:  "兔瓜",
		Age:   11,
		Class: "A1",
	}
	user2 := User{
		Id:    2,
		Name:  "",
		Age:   13,
		Class: "A1",
	}
	user3 := User{
		Id:    2,
		Name:  "兔罗",
		Age:   12,
		Class: "A2",
	}
	users = append(users, &user1, &user2, &user3)
	file, err := GenExcel([]string{"兔学院学生表"}, users)
	if err != nil {
		fmt.Printf("%v\n", err)
	}
	err = file.SaveAs("user.xlsx")
	if err != nil {
		fmt.Printf("%v\n", err)
	}
}
