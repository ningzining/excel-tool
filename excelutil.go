package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
)

// 获取excel的列编号
func genExcelField(num int) string {
	var (
		Str  string
		k    int
		temp []int
	)
	Slice := []string{"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
		"P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

	if num > 26 {
		for {
			k = num % 26
			if k == 0 {
				temp = append(temp, 26)
				k = 26
			} else {
				temp = append(temp, k)
			}
			num = (num - k) / 26
			if num <= 26 {
				temp = append(temp, num)
				break
			}
		}
	} else {
		return Slice[num]
	}

	for _, value := range temp {
		Str = Slice[value] + Str
	}
	return Str
}

// 获取结构体的tag中的excel列表以及tag对应的excel列
func getFields(v interface{}) (fields []string, fieldMap map[string]string, err error) {
	fieldMap = make(map[string]string, 8)
	var item interface{}
	switch reflect.TypeOf(v).Kind() {
	case reflect.Array, reflect.Slice:
		values := reflect.ValueOf(v)
		if values.Len() == 0 {
			return
		}
		item = values.Index(0).Interface()
	case reflect.Struct:
		reflect.ValueOf(v).Interface()
	default:
		err = errors.New(fmt.Sprintf("type %v not support", reflect.TypeOf(v).Kind()))
		return
	}
	typeOf := reflect.TypeOf(item)
	for i := 0; i < typeOf.NumField(); i++ {
		field := typeOf.Field(i).Tag.Get("excel")
		fields = append(fields, field)
		fieldMap[field] = genExcelField(len(fields))
	}
	return
}

// 获取结构体转为map[excelTag]structVal
func struct2MapList(v interface{}) (mapList []map[string]string) {
	switch reflect.TypeOf(v).Kind() {
	case reflect.Array, reflect.Slice:
		values := reflect.ValueOf(v)
		for i := 0; i < values.Len(); i++ {
			value := values.Index(i).Interface()
			typeOf := reflect.TypeOf(value)
			m := make(map[string]string)
			// 获取列所对应的值
			for j := 0; j < typeOf.NumField(); j++ {
				structField := typeOf.Field(j)
				name := reflect.ValueOf(value).FieldByName(structField.Name)
				field := typeOf.Field(j).Tag.Get("excel")
				m[field] = fmt.Sprintf("%v", name)
			}
			mapList = append(mapList, m)
		}
	}
	return
}

// 生成excel，row为起始行从1开始
func GenExcel(v interface{}, row int) (file *excelize.File, err error) {
	fields, fieldMap, err := getFields(v)
	if err != nil {
		return
	}
	file = excelize.NewFile()
	sheetName := "sheet1"
	for _, field := range fields {
		err = file.SetCellValue(sheetName, fmt.Sprintf("%s%d", fieldMap[field], row), field)
	}
	list := struct2MapList(v)
	for _, m := range list {
		row++
		for key, value := range m {
			err = file.SetCellValue(sheetName, fmt.Sprintf("%s%d", fieldMap[key], row), value)
		}
	}
	return
}
