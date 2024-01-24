package main

import (
	"errors"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"path"
	"reflect"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	tagKey         = "excel"
	ignoreTagValue = "-"

	cellTypeTag     = "cellType"
	cellTypePicture = "picture"
)

// SetRow 设置数据行
func (e *Excel) SetRow(data []string) *Excel {
	cellName, err := excelize.CoordinatesToCellName(1, e.Options.Row)
	if err != nil {
		e.Error = err
		return e
	}
	if err := e.File.SetSheetRow(e.Options.SheetName, cellName, &data); err != nil {
		e.Error = err
		return e
	}
	e.Options.Row++
	return e
}

// SetSliceDataWithTag 设置数据和tag行
func (e *Excel) SetSliceDataWithTag(slice any) *Excel {
	if err := e.checkSlice(slice); err != nil {
		e.Error = err
		return e
	}

	v := reflect.Indirect(reflect.ValueOf(slice))
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
	if err := e.checkSlice(slice); err != nil {
		e.Error = err
		return e
	}

	if err := e.setSheetRow(slice); err != nil {
		e.Error = err
		return e
	}
	return e
}

// 检查切片结构是否符合标准
func (e *Excel) checkSlice(slice any) error {
	sliceValue := reflect.ValueOf(slice)
	if sliceValue.IsNil() {
		return errors.New("切片为nil")
	}
	v := reflect.Indirect(sliceValue)
	if v.Type().Kind() != reflect.Slice && v.Type().Kind() != reflect.Array {
		return errors.New("目前只支持切片类型生成excel")
	}

	return nil
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
	cellName, err := excelize.CoordinatesToCellName(1, e.Options.Row)
	if err != nil {
		return err
	}
	if err := e.File.SetSheetRow(e.Options.SheetName, cellName, &headerRows); err != nil {
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

			cellType := structType.Field(j).Tag.Get(cellTypeTag)
			value := structValue.Field(j).Interface()
			if err := e.setCellValue(cellType, value, col); err != nil {
				return err
			}

			col++
		}

		e.Options.Row++
	}

	return nil
}

func (e *Excel) setCellValue(cellType string, value any, col int) error {
	cellName, err := excelize.CoordinatesToCellName(col, e.Options.Row)
	if err != nil {
		return err
	}

	switch cellType {
	case cellTypePicture: // 图片类型的值
		picPath, ok := value.(string)
		if !ok {
			return errors.New("picture tag value set error")
		}
		if picPath != "" {
			fileBytes, err := readBytesFromHttpUrl(picPath)
			if err != nil {
				return err
			}

			if err := e.File.AddPictureFromBytes(
				e.Options.SheetName,
				cellName,
				&excelize.Picture{
					Extension: strings.ToLower(path.Ext(picPath)),
					File:      fileBytes,
					Format:    nil,
				},
			); err != nil {
				return err
			}
		}
	default:
		if err := e.File.SetCellValue(e.Options.SheetName, cellName, value); err != nil {
			return err
		}
	}

	return nil
}
