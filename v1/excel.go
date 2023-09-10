package v1

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
)

const (
	startCol = "A"
)

type Excel struct {
	File      *excelize.File
	SheetName string
	Row       int
}

// Generate 生成excel
func Generate[T any](headers []string, slice []T) (file *excelize.File, err error) {
	sheetName := "sheet1"

	excel := &Excel{
		File:      excelize.NewFile(),
		SheetName: sheetName,
	}

	if err := setSheetHeaders(excel, headers); err != nil {
		return nil, err
	}

	if err := setSheetData[T](excel, slice); err != nil {
		return nil, err
	}

	return excel.File, nil
}

// SetSheetHeaders 设置表头
func setSheetHeaders(excel *Excel, headers []string) error {
	excel.Row++
	if err := excel.File.SetSheetRow(excel.SheetName, fmt.Sprintf("%s%d", startCol, excel.Row), &headers); err != nil {
		return err
	}
	return nil
}

// SetSheetData 设置数据
func setSheetData[T any](excel *Excel, data []T) (err error) {
	v := reflect.Indirect(reflect.ValueOf(data))
	if v.Type().Kind() != reflect.Slice {
		return errors.New("目前只支持切片类型生成excel")
	}
	if v.Len() == 0 {
		return
	}

	rowNum := 2
	if err := setSheetTitle[T](excel, data[0]); err != nil {
		return err
	} else {
		rowNum++
	}

	if err := setSheetRowData[T](excel, data); err != nil {
		return err
	}
	return
}

// 设置行数据
func setSheetRowData[T any](excel *Excel, data []T) error {
	for i := 0; i < len(data); i++ {
		structValue := reflect.Indirect(reflect.ValueOf(data[i]))
		structType := reflect.TypeOf(data[i])
		if structType.Kind() == reflect.Ptr {
			structType = structType.Elem()
		}

		var rowData []any
		for j := 0; j < structType.NumField(); j++ {
			value := structValue.FieldByName(structType.Field(j).Name)
			rowData = append(rowData, value.Interface())
		}

		excel.Row++
		err := excel.File.SetSheetRow(excel.SheetName, fmt.Sprintf("%s%d", startCol, excel.Row), &rowData)
		if err != nil {
			return err
		}
	}

	return nil
}

// 设置sheet标题
func setSheetTitle[T any](excel *Excel, data T) error {
	var titles []string
	structType := reflect.TypeOf(data)
	if structType.Kind() == reflect.Ptr {
		structType = structType.Elem()
	}
	for i := 0; i < structType.NumField(); i++ {
		field := structType.Field(i).Tag.Get("excel")
		titles = append(titles, field)
	}
	excel.Row++
	if err := excel.File.SetSheetRow(excel.SheetName, fmt.Sprintf("%s%d", startCol, excel.Row), &titles); err != nil {
		return err
	}

	return nil
}
