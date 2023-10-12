package api

import (
	"fmt"
	"math"
	"strconv"
	"unicode"
)

type CellPosition struct {
	RowIndex    int
	ColumnIndex int
}

func Origin() *CellPosition {
	return &CellPosition{1, 1}
}

func (c *CellPosition) ToAlphaNumeric() (string, error) {
	col, err := indexToColumn(c.ColumnIndex)
	if err != nil {
		return "", err
	}
	row := strconv.Itoa(c.RowIndex)
	return col + row, nil
}

func FromAlphaNumric(alphaNumeric string) (*CellPosition, error) {
	splitIndex := 0
	for i, r := range alphaNumeric {
		if unicode.IsDigit(r) {
			splitIndex = i
			break
		}
	}
	col, err := columnToIndex(alphaNumeric[:splitIndex])
	if err != nil {
		return nil, err
	}
	row, err := strconv.Atoi(alphaNumeric[splitIndex:])
	if err != nil {
		return nil, err
	}

	return &CellPosition{
		RowIndex:    row,
		ColumnIndex: col,
	}, nil
}

func (c *CellPosition) Offset(rowOffset int, columnOffset int) *CellPosition {
	return &CellPosition{
		RowIndex:    c.RowIndex + rowOffset,
		ColumnIndex: c.ColumnIndex + columnOffset,
	}
}

// indexToColumn takes in an index value & converts it to A1 Notation
// Index 1 is Column A
// E.g. 3 == C, 29 == AC, 731 == ABC
func indexToColumn(index int) (string, error) {

	// Validate index size
	maxIndex := 18278
	if index > maxIndex {
		return "", fmt.Errorf("index cannot be greater than %v (column ZZZ)", maxIndex)
	}

	// Get column from index
	l := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if index > 26 {
		letterA, _ := indexToColumn(int(math.Floor(float64(index-1) / 26)))
		letterB, _ := indexToColumn(index % 26)
		return letterA + letterB, nil
	} else {
		if index == 0 {
			index = 26
		}
		return string(l[index-1]), nil
	}

}

//https://stackoverflow.com/questions/70806630/convert-index-to-column-a1-notation-and-vice-versa
// columnToIndex takes in A1 Notation & converts it to an index value
// Column A is index 1
// E.g. C == 3, AC == 29, ABC == 731
func columnToIndex(column string) (int, error) {

	// Calculate index from column string
	var index int
	var a uint8 = "A"[0]
	var z uint8 = "Z"[0]
	var alphabet = z - a + 1
	i := 1
	for n := len(column) - 1; n >= 0; n-- {
		r := column[n]
		if r < a || r > z {
			return 0, fmt.Errorf("invalid character in column, expected A-Z but got [%c]", r)
		}
		runePos := int(r-a) + 1
		index += runePos * int(math.Pow(float64(alphabet), float64(i-1)))
		i++
	}

	// Return column index & success
	return index, nil

}
