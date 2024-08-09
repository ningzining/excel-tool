package main

import (
	"io"

	"github.com/xuri/excelize/v2"
)

type Excel struct {
	File    *excelize.File
	Options *options
	Error   error
}

func New(opts ...Option) *Excel {
	o := &options{
		Row: 1,
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

func (e *Excel) Write(w io.Writer) *Excel {
	if err := e.File.Write(w); err != nil {
		e.Error = err
	}
	return e
}

func (e *Excel) SaveAs(name string) *Excel {
	if err := e.File.SaveAs(name); err != nil {
		e.Error = err
	}
	return e
}
