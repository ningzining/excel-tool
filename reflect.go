package main

import (
	"reflect"
)

func typeOfStruct(data any) reflect.Type {
	target := reflect.TypeOf(data)
	if target.Kind() != reflect.Pointer {
		return target
	}

	return target.Elem()
}

func valueOfStruct(data any) reflect.Value {
	return reflect.Indirect(reflect.ValueOf(data))
}
