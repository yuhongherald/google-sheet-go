package api

import "google.golang.org/api/sheets/v4"

// Color values from 0 to 1
type Color struct {
	R float64
	G float64
	B float64
	A float64
}

func Lerp(color1 *Color, color2 *Color, alpha float64) *Color {
	oneMinusAlpha := 1 - alpha

	return &Color{
		R: color1.R*oneMinusAlpha + color2.R*alpha,
		G: color1.G*oneMinusAlpha + color2.G*alpha,
		B: color1.B*oneMinusAlpha + color2.B*alpha,
		A: color1.A*oneMinusAlpha + color2.A*alpha,
	}
}

func (c *Color) toSheetsColor() *sheets.Color {
	return &sheets.Color{
		Red:   c.R,
		Green: c.G,
		Blue:  c.B,
		Alpha: c.A,
	}
}
