package api

import (
	"encoding/json"
	"io"
	"os"
)

type Chart struct {
	Title       string `json:"title"`
	TopLeft     Coord  `json:"top_left"`
	Size        Size   `json:"size"`
	XAxisTitle  string `json:"x_axis_title"`
	YAxisTitle  string `json:"y_axis_title"`
	LabelColumn string `json:"label_column"`
	DataColumn  string `json:"data_column"`
}

// 0-indexed coordinate
type Coord struct {
	X int64 `json:"x"`
	Y int64 `json:"y"`
}

type Size struct {
	Height int64 `json:"height"`
	Width  int64 `json:"width"`
}

func ReadFromFile(filename string) ([]*Chart, error) {
	jsonFile, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer jsonFile.Close()

	b, err := io.ReadAll(jsonFile)
	if err != nil {
		return nil, err
	}

	charts := make([]*Chart, 0)
	err = json.Unmarshal(b, &charts)
	return charts, err
}
