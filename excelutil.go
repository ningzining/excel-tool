package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
)

// GenExcel 生成excel
func GenExcel(headers []string, slice interface{}) (file *excelize.File, err error) {
	file = excelize.NewFile()
	sheetName := "sheet1"

	if err := setSheetHeaders(file, sheetName, headers); err != nil {
		return nil, err
	}

	if err := setSheetData(file, sheetName, slice); err != nil {
		return nil, err
	}

	return
}

// 设置表头
func setSheetHeaders(file *excelize.File, sheetName string, headers []string) error {
	if err := file.SetSheetRow(sheetName, "A1", &headers); err != nil {
		return err
	}
	return nil
}

// 设置数据
func setSheetData(file *excelize.File, sheetName string, data any) (err error) {
	v := reflect.Indirect(reflect.ValueOf(data))
	if v.Type().Kind() != reflect.Slice {
		return errors.New("目前只支持切片类型生成excel")
	}
	if v.Len() == 0 {
		return
	}
	structValue := reflect.Indirect(v.Index(0)).Interface()
	rowNum := 2
	if err := setSheetTitle(file, sheetName, rowNum, structValue); err != nil {
		return err
	} else {
		rowNum++
	}

	if err := setSheetRowData(file, sheetName, rowNum, v); err != nil {
		return err
	}
	return
}

// 设置行数据
func setSheetRowData(file *excelize.File, sheetName string, rowNum int, data reflect.Value) error {
	for i := 0; i < data.Len(); i++ {
		structValue := reflect.Indirect(data.Index(i)).Interface()
		structType := reflect.TypeOf(reflect.Indirect(data.Index(i)).Interface())

		var rowData []any
		for j := 0; j < structType.NumField(); j++ {
			value := reflect.ValueOf(structValue).FieldByName(structType.Field(j).Name)
			rowData = append(rowData, value.Interface())
		}

		err := file.SetSheetRow(sheetName, fmt.Sprintf("A%d", rowNum), &rowData)
		if err != nil {
			return err
		}
		rowNum++
	}

	return nil
}

// 设置sheet标题
func setSheetTitle(file *excelize.File, sheetName string, rowNum int, v any) error {
	var titles []string
	itemType := reflect.TypeOf(v)
	for i := 0; i < itemType.NumField(); i++ {
		field := itemType.Field(i).Tag.Get("excel")
		titles = append(titles, field)
	}
	if err := file.SetSheetRow(sheetName, fmt.Sprintf("A%d", rowNum), &titles); err != nil {
		return err
	}

	return nil
}
