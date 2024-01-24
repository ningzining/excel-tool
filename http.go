package main

import (
	"io"
	"net/http"
)

func readBytesFromHttpUrl(url string) ([]byte, error) {
	response, err := http.Get(url)
	if err != nil {
		return nil, err
	}
	defer response.Body.Close()

	bytes, err := io.ReadAll(response.Body)
	if err != nil {
		return nil, err
	}

	return bytes, nil
}
