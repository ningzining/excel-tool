package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"reflect"
)

const (
	startCol         = "A"
	defaultSheetName = "Sheet1"
)

type Excel struct {
	File    *excelize.File
	Options *options
	Error   error
}

func New(opts ...Option) *Excel {
	o := &options{
		SheetName: defaultSheetName,
		Row:       1,
	}
	for _, opt := range opts {
		opt.apply(o)
	}
	excel := &Excel{
		File:    excelize.NewFile(),
		Options: o,
	}
	return excel
}

func (e *Excel) SaveAs(name string) *Excel {
	if err := e.File.SaveAs(name); err != nil {
		e.Error = err
	}
	return e
}

// SetTitles 设置表头
func (e *Excel) SetTitles(titles []string) *Excel {
	if err := e.File.SetSheetRow(e.Options.SheetName, fmt.Sprintf("%s%d", startCol, e.Options.Row), &titles); err != nil {
		e.Error = err
	}
	e.Options.Row++
	return e
}

func (e *Excel) SetData(slice any) *Excel {
	sliceValue := reflect.ValueOf(slice)
	if sliceValue.IsNil() {
		return e
	}
	v := reflect.Indirect(sliceValue)
	if v.Type().Kind() != reflect.Slice {
		e.Error = errors.New("目前只支持切片类型生成excel")
		return e
	}
	if v.Len() == 0 {
		return e
	}

	if err := e.setSheetHeaderRow(v.Index(0).Interface()); err != nil {
		e.Error = err
		return e
	}
	if err := e.setSheetRow(slice); err != nil {
		e.Error = err
		return e
	}
	return e
}

// 设置headerRow标题行
func (e *Excel) setSheetHeaderRow(data any) error {
	var headerRows []string
	structType := reflect.TypeOf(data)
	if structType.Kind() == reflect.Ptr {
		structType = structType.Elem()
	}
	for i := 0; i < structType.NumField(); i++ {
		field := structType.Field(i).Tag.Get("excel")
		headerRows = append(headerRows, field)
	}

	if err := e.File.SetSheetRow(e.Options.SheetName, fmt.Sprintf("%s%d", startCol, e.Options.Row), &headerRows); err != nil {
		return err
	}
	e.Options.Row++

	return nil
}

func (e *Excel) setSheetRow(slice any) error {
	v := reflect.Indirect(reflect.ValueOf(slice))
	for i := 0; i < v.Len(); i++ {
		structValue := reflect.Indirect(v.Index(i))
		structType := reflect.TypeOf(v.Index(i).Interface())
		if structType.Kind() == reflect.Ptr {
			structType = structType.Elem()
		}

		var rows []any
		for j := 0; j < structType.NumField(); j++ {
			value := structValue.Field(j).Interface()
			rows = append(rows, value)
		}

		err := e.File.SetSheetRow(e.Options.SheetName, fmt.Sprintf("%s%d", startCol, e.Options.Row), &rows)
		if err != nil {
			return err
		}
		e.Options.Row++
	}

	return nil
}

func (e *Excel) Write(w io.Writer) *Excel {
	if err := e.File.Write(w); err != nil {
		e.Error = err
	}
	return e
}
