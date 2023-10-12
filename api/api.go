package api

import (
	"context"
	"encoding/base64"
	"errors"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/gmail/v1"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"math"
)

type Service struct {
	sheets *sheets.Service
	drive  *drive.Service
	gmail  *gmail.Service
}

func NewService(ctx context.Context, credentialsJson []byte) (*Service, error) {
	sheetsService, err := sheets.NewService(ctx, option.WithCredentialsJSON(credentialsJson))
	if err != nil {
		return nil, err
	}

	driveService, err := drive.NewService(ctx, option.WithCredentialsJSON(credentialsJson))
	if err != nil {
		return nil, err
	}

	gmailService, err := gmail.NewService(ctx, option.WithCredentialsJSON(credentialsJson))
	if err != nil {
		return nil, err
	}

	return &Service{
		sheets: sheetsService,
		drive:  driveService,
		gmail:  gmailService,
	}, nil
}

func (s *Service) Share(fileId string, email string) error {
	permission := &drive.Permission{
		EmailAddress: email,
		Role:         "writer",
		Type:         "user",
	}
	_, err := s.drive.Permissions.Create(fileId, permission).Do()
	return err
}

func (s *Service) Create(title string) (string, error) {
	spreadsheet, err := s.sheets.Spreadsheets.Create(&sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: title,
		},
	}).Do()
	if err != nil {
		return "", err
	}
	return spreadsheet.SpreadsheetId, nil
}

func (s *Service) Recreate(spreadsheetId string, title string) error {
	resp, err := s.sheets.Spreadsheets.Get(spreadsheetId).IncludeGridData(true).Do()
	if err != nil {
		return err
	}

	if len(resp.Sheets) == 0 {
		return errors.New("0 sheets")
	}

	var requests []*sheets.Request
	if title != "" {
		requests = append(requests, &sheets.Request{
			UpdateSpreadsheetProperties: &sheets.UpdateSpreadsheetPropertiesRequest{
				Fields: "title",
				Properties: &sheets.SpreadsheetProperties{
					Title: title,
				},
			},
		})
	}

	hasSheet1 := false
	for _, sheet := range resp.Sheets {
		if sheet.Properties.Title == "Sheet1" {
			requests = append(requests, []*sheets.Request{
				{
					UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
						Fields: "title",
						Properties: &sheets.SheetProperties{
							SheetId: sheet.Properties.SheetId,
							Title:   "Temp",
						},
					},
				},
				{
					AddSheet: &sheets.AddSheetRequest{
						Properties: &sheets.SheetProperties{
							Title: "Sheet1",
						},
					},
				},
				{
					DeleteSheet: &sheets.DeleteSheetRequest{
						SheetId: sheet.Properties.SheetId,
					},
				},
			}...)
			hasSheet1 = true
			break
		}
	}

	if !hasSheet1 {
		requests = append(requests, &sheets.Request{
			AddSheet: &sheets.AddSheetRequest{
				Properties: &sheets.SheetProperties{
					Title: "Sheet1",
				},
			},
		})
	}

	_, err = s.sheets.Spreadsheets.BatchUpdate(spreadsheetId, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}).Do()
	return err
}

func (s *Service) InsertTable(spreadsheetId string, cellPosition *CellPosition, table [][]string) error {
	if len(table) == 0 || len(table[0]) == 0 {
		return errors.New("Attempting to insert empty table")
	}
	height := len(table)
	width := len(table[0])

	var tableRaw [][]interface{}
	for _, row := range table {
		var tableRawRow []interface{}
		for _, column := range row {
			tableRawRow = append(tableRawRow, column)
		}
		tableRaw = append(tableRaw, tableRawRow)
	}
	start, err := cellPosition.ToAlphaNumeric()
	if err != nil {
		return err
	}

	end, err := cellPosition.Offset(height-1, width-1).ToAlphaNumeric()
	if err != nil {
		return err
	}

	_, err = s.sheets.Spreadsheets.Values.Update(spreadsheetId, start+":"+end, &sheets.ValueRange{
		Values: tableRaw,
	}).ValueInputOption("USER_ENTERED").Do()
	return err
}

func (s *Service) AddChart(spreadsheetId string, chart *Chart) error {
	resp, err := s.sheets.Spreadsheets.Get(spreadsheetId).IncludeGridData(true).Do()
	if err != nil {
		return err
	}
	if len(resp.Sheets) == 0 {
		return errors.New("0 sheets")
	}
	sheetId := resp.Sheets[0].Properties.SheetId

	labelIndex := int64(0)
	dataIndex := int64(0)
	startColumn := resp.Sheets[0].Data[0].StartColumn
	startRow := resp.Sheets[0].Data[0].StartRow
	endRow := int64(len(resp.Sheets[0].Data[0].RowData)) + startRow
	for index, headerCell := range resp.Sheets[0].Data[0].RowData[0].Values {
		if headerCell.EffectiveValue.StringValue == nil {
			continue
		}
		header := *headerCell.EffectiveValue.StringValue
		if header == chart.LabelColumn {
			labelIndex = int64(index)
		}
		if header == chart.DataColumn {
			dataIndex = int64(index)
		}
	}

	hasData := false
	max := -math.MaxFloat64
	min := math.MaxFloat64
	for _, row := range resp.Sheets[0].Data[0].RowData {
		numberValue := row.Values[dataIndex].EffectiveValue.NumberValue
		if numberValue != nil {
			hasData = true
			max = math.Max(max, *numberValue)
			min = math.Min(min, *numberValue)
		}
	}
	var viewWindowOptions *sheets.ChartAxisViewWindowOptions
	if hasData {
		viewWindowOptions = &sheets.ChartAxisViewWindowOptions{
			ViewWindowMin: min,
			ViewWindowMax: max,
		}
	}

	request := &sheets.Request{
		AddChart: &sheets.AddChartRequest{
			Chart: &sheets.EmbeddedChart{
				Position: &sheets.EmbeddedObjectPosition{
					OverlayPosition: &sheets.OverlayPosition{
						AnchorCell: &sheets.GridCoordinate{
							ColumnIndex: 0,
							RowIndex:    0,
							SheetId:     sheetId,
						},
						HeightPixels:  chart.Size.Height,
						OffsetXPixels: chart.TopLeft.X,
						OffsetYPixels: chart.TopLeft.Y,
						WidthPixels:   chart.Size.Width,
					},
				},
				Spec: &sheets.ChartSpec{
					BasicChart: &sheets.BasicChartSpec{
						Axis: []*sheets.BasicChartAxis{
							{
								Format: &sheets.TextFormat{
									FontFamily: "Roboto",
								},
								Position:          "BOTTOM_AXIS",
								Title:             chart.XAxisTitle, // "tag",
								ViewWindowOptions: &sheets.ChartAxisViewWindowOptions{},
							},
							{
								Format: &sheets.TextFormat{
									FontFamily: "Roboto",
								},
								Position:          "LEFT_AXIS",
								Title:             chart.YAxisTitle,
								ViewWindowOptions: viewWindowOptions,
							},
						},
						ChartType: "LINE",
						Domains: []*sheets.BasicChartDomain{
							{
								Domain: &sheets.ChartData{
									SourceRange: &sheets.ChartSourceRange{
										Sources: []*sheets.GridRange{
											{
												StartColumnIndex: labelIndex + startColumn,
												EndColumnIndex:   labelIndex + startColumn + 1,
												StartRowIndex:    startRow,
												EndRowIndex:      endRow,
												SheetId:          sheetId,
											},
										},
									},
								},
							},
						},
						HeaderCount: 1,
						Series: []*sheets.BasicChartSeries{
							{
								DataLabel: &sheets.DataLabel{
									TextFormat: &sheets.TextFormat{
										FontFamily: "Roboto",
									},
									Type: "NONE",
								},
								Series: &sheets.ChartData{
									SourceRange: &sheets.ChartSourceRange{
										Sources: []*sheets.GridRange{
											{
												StartColumnIndex: dataIndex + startColumn,
												EndColumnIndex:   dataIndex + startColumn + 1,
												StartRowIndex:    startRow,
												EndRowIndex:      endRow,
												SheetId:          sheetId,
											},
										},
									},
								},
								TargetAxis: "LEFT_AXIS",
							},
						},
					},
					FontName:                "Roboto",
					HiddenDimensionStrategy: "SKIP_HIDDEN_ROWS_AND_COLUMNS",
					Title:                   chart.Title,
					TitleTextFormat: &sheets.TextFormat{
						FontFamily: "Roboto",
					},
				},
			},
		},
	}

	_, err = s.sheets.Spreadsheets.BatchUpdate(spreadsheetId, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			request,
		},
	}).Do()

	return err
}

func (s *Service) GetFirstSheetId(spreadsheetId string) (int64, error) {
	resp, err := s.sheets.Spreadsheets.Get(spreadsheetId).Do()
	if err != nil {
		return 0, err
	}
	if len(resp.Sheets) == 0 {
		return 0, errors.New("0 sheets")
	}
	return resp.Sheets[0].Properties.SheetId, nil
}

func (s *Service) Highlight(spreadsheetId string,
	startPosition *CellPosition, endPosition *CellPosition,
	lowerRange *Boundary, upperRange *Boundary,
	chosenColor *Color) error {

	resp, err := s.sheets.Spreadsheets.Get(spreadsheetId).IncludeGridData(true).Do()
	if err != nil {
		return err
	}
	if len(resp.Sheets) == 0 {
		return errors.New("0 sheets")
	}
	sheetId := resp.Sheets[0].Properties.SheetId

	var rows []*sheets.RowData
	for i := startPosition.RowIndex; i <= endPosition.RowIndex; i++ {
		var cols []*sheets.CellData
		for j := startPosition.ColumnIndex; j <= endPosition.ColumnIndex; j++ {
			rawValue := resp.Sheets[0].Data[0].RowData[int64(i)-resp.Sheets[0].Data[0].StartRow-1].Values[int64(j)-resp.Sheets[0].Data[0].StartColumn-1].EffectiveValue.NumberValue
			var value float64
			var hasValue bool

			if rawValue != nil {
				hasValue = true
				value = *rawValue
			}
			alpha := NormalizeBoundary(lowerRange, upperRange, value)

			var color *sheets.Color
			if hasValue && alpha >= 0 && alpha <= 1 {
				color = chosenColor.toSheetsColor()
			} else {
				// get the original color
				color = resp.Sheets[0].Data[0].RowData[int64(i)-resp.Sheets[0].Data[0].StartRow-1].Values[int64(j)-resp.Sheets[0].Data[0].StartColumn-1].EffectiveFormat.BackgroundColor
			}

			cols = append(cols, &sheets.CellData{
				UserEnteredFormat: &sheets.CellFormat{
					BackgroundColor: color,
				},
			})
		}
		rows = append(rows, &sheets.RowData{
			Values: cols,
		})
	}

	request := &sheets.Request{
		UpdateCells: &sheets.UpdateCellsRequest{
			Fields: "user_entered_format.background_color",
			Range: &sheets.GridRange{
				EndColumnIndex:   int64(endPosition.ColumnIndex),
				EndRowIndex:      int64(endPosition.RowIndex),
				SheetId:          sheetId,
				StartColumnIndex: int64(startPosition.ColumnIndex - 1),
				StartRowIndex:    int64(startPosition.RowIndex - 1),
			},
			Rows: rows,
		},
	}

	_, err = s.sheets.Spreadsheets.BatchUpdate(spreadsheetId, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			request,
		},
	}).Do()

	return err
}

func (s *Service) GradientHighlight(spreadsheetId string,
	startPosition *CellPosition, endPosition *CellPosition,
	lowerRange *Boundary, upperRange *Boundary,
	color1 *Color, color2 *Color) error {

	if lowerRange == upperRange {
		return errors.New("lower range equal upper range")
	}

	resp, err := s.sheets.Spreadsheets.Get(spreadsheetId).IncludeGridData(true).Do()
	if err != nil {
		return err
	}
	if len(resp.Sheets) == 0 {
		return errors.New("0 sheets")
	}
	sheetId := resp.Sheets[0].Properties.SheetId

	var rows []*sheets.RowData
	for i := startPosition.RowIndex; i <= endPosition.RowIndex; i++ {
		var cols []*sheets.CellData
		for j := startPosition.ColumnIndex; j <= endPosition.ColumnIndex; j++ {
			rawValue := resp.Sheets[0].Data[0].RowData[int64(i)-resp.Sheets[0].Data[0].StartRow-1].Values[int64(j)-resp.Sheets[0].Data[0].StartColumn-1].EffectiveValue.NumberValue
			var value float64
			var hasValue bool

			if rawValue != nil {
				hasValue = true
				value = *rawValue
			}
			alpha := NormalizeBoundary(lowerRange, upperRange, value)

			var color *sheets.Color
			if hasValue && alpha >= 0 && alpha <= 1 {
				color = Lerp(color1, color2, alpha).toSheetsColor()
			} else {
				// get the original color
				color = resp.Sheets[0].Data[0].RowData[int64(i)-resp.Sheets[0].Data[0].StartRow-1].Values[int64(j)-resp.Sheets[0].Data[0].StartColumn-1].EffectiveFormat.BackgroundColor
			}

			cols = append(cols, &sheets.CellData{
				UserEnteredFormat: &sheets.CellFormat{
					BackgroundColor: color,
				},
			})
		}
		rows = append(rows, &sheets.RowData{
			Values: cols,
		})
	}

	request := &sheets.Request{
		UpdateCells: &sheets.UpdateCellsRequest{
			Fields: "user_entered_format.background_color",
			Range: &sheets.GridRange{
				EndColumnIndex:   int64(endPosition.ColumnIndex),
				EndRowIndex:      int64(endPosition.RowIndex),
				SheetId:          sheetId,
				StartColumnIndex: int64(startPosition.ColumnIndex - 1),
				StartRowIndex:    int64(startPosition.RowIndex - 1),
			},
			Rows: rows,
		},
	}

	_, err = s.sheets.Spreadsheets.BatchUpdate(spreadsheetId, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			request,
		},
	}).Do()

	return err
}

func (s *Service) SendEmail(users string, title string, message string) error {
	messageObject := &gmail.Message{
		Payload: &gmail.MessagePart{
			Body: &gmail.MessagePartBody{
				Data: base64.StdEncoding.EncodeToString([]byte(message)),
			},
			Headers: []*gmail.MessagePartHeader{
				{
					Name:  "To",
					Value: users,
				},
				{
					Name:  "Subject",
					Value: title,
				},
			},
			MimeType: "text/html",
		},
	}

	_, err := s.gmail.Users.Messages.Send("me", messageObject).Do()
	return err
}
