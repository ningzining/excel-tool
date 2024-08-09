package main

type options struct {
	Row int
}

type Option interface {
	apply(o *options)
}

type optionFunc func(o *options)

func (of optionFunc) apply(o *options) {
	of(o)
}

func WithStartRow(row int) Option {
	return optionFunc(func(o *options) {
		o.Row = row
	})
}
