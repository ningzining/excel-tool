package main

import (
	"os"
	"testing"
)

type User struct {
	Id          int    `excel:"编号" json:"id"`
	Name        string `excel:"学生名字" json:"name"`
	Age         int    `excel:"年龄" json:"age"`
	Class       string `excel:"班级" json:"class"`
	IgnoreClass string `excel:"-"`
	Pic         string `excel:"图片" cellType:"picture"`
}

func TestSaveExcel(t *testing.T) {
	users := mockUsers()
	if err := New().
		SetRow([]string{"兔学院学生表"}).
		SetSliceDataWithTag(&users).
		SaveAs("sheet_name.xlsx").
		Error; err != nil {
		t.Error(err)
		return
	}
}

func TestSaveNilExcel(t *testing.T) {
	var users []*User
	if err := New().
		SetRow([]string{"兔学院学生表"}).
		SetSliceDataWithTag(users).
		SaveAs("user_generate_nil.xlsx").
		Error; err != nil {
		t.Error(err)
		return
	}
}

func TestWriterExcel(t *testing.T) {
	file, err := os.Create("user_write.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	defer file.Close()

	users := mockUsers()
	if err := New().
		SetRow([]string{"兔学院学生表"}).
		SetSliceDataWithTag(users).
		Write(file).
		Error; err != nil {
		t.Error(err)
		return
	}
}

func mockUsers() []*User {
	var users []*User
	user1 := User{
		Id:    1,
		Name:  "兔瓜",
		Age:   11,
		Class: "A1",
		Pic:   "",
	}
	user2 := User{
		Id:    2,
		Name:  "兔柠",
		Age:   13,
		Class: "A1",
		Pic:   "",
	}
	user3 := User{
		Id:    2,
		Name:  "兔罗",
		Age:   12,
		Class: "A2",
		Pic:   "",
	}
	users = append(users, &user1, &user2, &user3)
	return users
}
