package main

type Config struct {
	SheetName string
	Row       int
}

type OptionFunc func(config *Config)

func WithSheetName(sheetName string) OptionFunc {
	return func(config *Config) {
		config.SheetName = sheetName
	}
}
