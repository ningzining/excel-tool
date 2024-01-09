package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"net/http"
	"path"
	"reflect"
	"strings"
)

const (
	startCol         = "A"
	defaultSheetName = "Sheet1"
	tagKey           = "excel"
	ignoreTagValue   = "-"
	cellType         = "cellType"
	picCell          = "picture"
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

// SetRow 设置数据行
func (e *Excel) SetRow(data []string) *Excel {
	if err := e.File.SetSheetRow(e.Options.SheetName, fmt.Sprintf("%s%d", startCol, e.Options.Row), &data); err != nil {
		e.Error = err
	}
	e.Options.Row++
	return e
}

// SetSliceDataWithTag 设置数据和tag行
func (e *Excel) SetSliceDataWithTag(slice any) *Excel {
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

// SetSliceData 设置数据行
func (e *Excel) SetSliceData(slice any) *Excel {
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
		field := structType.Field(i).Tag.Get(tagKey)
		if field == ignoreTagValue {
			continue
		}
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

		col := 1
		for j := 0; j < structType.NumField(); j++ {
			tag := structType.Field(j).Tag.Get(tagKey)
			if tag == ignoreTagValue {
				continue
			}
			cellName, err := excelize.CoordinatesToCellName(col, e.Options.Row)
			if err != nil {
				return err
			}
			value := structValue.Field(j).Interface()
			if structType.Field(j).Tag.Get(cellType) == picCell {
				picPath := value.(string)
				if picPath != "" {
					ext := strings.ToLower(path.Ext(picPath))
					response, err := http.Get(picPath)
					if err != nil {
						return err
					}
					defer response.Body.Close()

					bytes, err := io.ReadAll(response.Body)
					if err != nil {
						return err
					}
					if err := e.File.AddPictureFromBytes(e.Options.SheetName, cellName, &excelize.Picture{Extension: ext, File: bytes, Format: nil}); err != nil {
						return err
					}
				}
			} else {
				if err := e.File.SetCellValue(e.Options.SheetName, cellName, value); err != nil {
					return err
				}
			}
			col++
			//if structType.Field(j).Tag.Get(cellType) == picCell {
			//	picCol = j
			//	picPath = structValue.Field(j).Interface().(string)
			//	if picPath == "" {
			//		continue
			//	}
			//	cellName, err := excelize.CoordinatesToCellName(picCol, e.Options.Row)
			//	if err != nil {
			//		return err
			//	}
			//	if err := e.File.AddPicture(e.Options.SheetName, cellName, picPath, nil); err != nil {
			//		return err
			//	}
			//} else {
			//}

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
