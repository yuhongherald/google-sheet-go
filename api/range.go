package api

import "math"

type Boundary struct {
	Value float64
	Clamp bool
}

func NormalizeBoundary(lower *Boundary, upper *Boundary, value float64) float64 {
	result := (value - lower.Value) / (upper.Value - lower.Value)
	if lower.Clamp {
		result = math.Max(result, 0)
	}
	if upper.Clamp {
		result = math.Min(result, 1)
	}
	return result
}

func Normalize(lower float64, upper float64, value float64) float64 {
	return (value - lower) / (upper - lower)
}
