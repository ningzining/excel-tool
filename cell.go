package main

import (
	"errors"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"path"
	"reflect"

	"github.com/xuri/excelize/v2"
)

const (
	tagKey         = "excel"
	ignoreTagValue = "-"

	cellTypeTag     = "cellType"
	cellTypePicture = "picture"
	cellTypeLink    = "link"
)

// SetSheetRow 设置数据行
func (e *Excel) SetSheetRow(sheetName string, row []string) *Excel {
	cellName, err := excelize.CoordinatesToCellName(1, e.Options.Row)
	if err != nil {
		e.Error = err
		return e
	}

	if err := e.File.SetSheetRow(sheetName, cellName, &row); err != nil {
		e.Error = err
		return e
	}
	e.Options.Row++

	return e
}

// SetSheetRowsWithHeader 设置数据和tag行
func (e *Excel) SetSheetRowsWithHeader(sheetName string, slice interface{}) *Excel {
	if err := e.checkSlice(slice); err != nil {
		e.Error = err
		return e
	}

	v := reflect.Indirect(reflect.ValueOf(slice))
	if v.Len() == 0 {
		return e
	}
	if err := e.setSheetHeaderRow(sheetName, v.Index(0).Interface()); err != nil {
		e.Error = err
		return e
	}

	if err := e.setSheetRow(sheetName, slice); err != nil {
		e.Error = err
		return e
	}
	return e
}

// 检查切片结构是否符合标准
func (e *Excel) checkSlice(slice interface{}) error {
	if slice == nil {
		return errors.New("切片为nil")
	}

	v := reflect.Indirect(reflect.ValueOf(slice))
	if v.Type().Kind() != reflect.Slice && v.Type().Kind() != reflect.Array {
		return errors.New("目前只支持切片类型生成excel")
	}

	return nil
}

// 设置headerRow标题行
func (e *Excel) setSheetHeaderRow(sheetName string, data any) error {
	var headers []string
	structType := typeOfStruct(data)

	for i := 0; i < structType.NumField(); i++ {
		field := structType.Field(i).Tag.Get(tagKey)
		if field == ignoreTagValue {
			continue
		}
		headers = append(headers, field)
	}
	cellName, err := excelize.CoordinatesToCellName(1, e.Options.Row)
	if err != nil {
		return err
	}

	if err := e.File.SetSheetRow(sheetName, cellName, &headers); err != nil {
		return err
	}
	e.Options.Row++

	return nil
}

func (e *Excel) setSheetRow(sheetName string, slice interface{}) error {
	v := reflect.Indirect(reflect.ValueOf(slice))
	for i := 0; i < v.Len(); i++ {
		structValue := valueOfStruct(v.Index(i).Interface())
		structType := typeOfStruct(v.Index(i).Interface())

		col := 1
		for j := 0; j < structType.NumField(); j++ {
			tag := structType.Field(j).Tag.Get(tagKey)
			if tag == ignoreTagValue {
				continue
			}

			cellType := structType.Field(j).Tag.Get(cellTypeTag)
			value := structValue.Field(j).Interface()
			if err := e.setCellValue(sheetName, cellType, value, col); err != nil {
				return err
			}

			col++
		}

		e.Options.Row++
	}

	return nil
}

func (e *Excel) setCellValue(sheetName string, cellType string, value any, col int) error {
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
			ext := path.Ext(picPath)
			if err := e.File.AddPictureFromBytes(
				sheetName,
				cellName,
				&excelize.Picture{
					Extension: ext,
					File:      fileBytes,
					Format: &excelize.GraphicOptions{
						LockAspectRatio: true,
						AutoFit:         true,
					},
				},
			); err != nil {
				return err
			}
		}
	case cellTypeLink:
		if err := e.File.SetCellHyperLink(sheetName, cellName, value.(string), "External"); err != nil {
			return err
		}
	default:
		if err := e.File.SetCellValue(sheetName, cellName, value); err != nil {
			return err
		}
	}

	return nil
}

func (e *Excel) UpdateLinkedValue() *Excel {
	if err := e.File.UpdateLinkedValue(); err != nil {
		e.Error = err
		return e
	}

	return e
}
