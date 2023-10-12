package main

import (
	"context"
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"github.com/yuhongherald/google-sheet-go/api"
	"log"
	"math"
	"os"
	"sort"
	"strconv"
	"strings"
)

func main() {
	titleP := flag.String("title", "title", "Title of Google Sheet")
	googleCredentialsP := flag.String("google-credentials", "{}", "Google credentials JSON string")
	contentFileP := flag.String("content-file", "", "Table contents to be uploaded, in csv format")
	chartFileP := flag.String("chart-file", "chart.json", "list of graph config objects")
	usersP := flag.String("users", "", "comma separated emails")
	sendEmailMessageP := flag.String("send-email-message", "", "Sender email. Leave this blank to not send email")
	highlightColumnsP := flag.String("highlight-columns", "", "Comma separated column names")
	spreadsheetIdP := flag.String("google-sheet-id", "", "Google Sheet id of existing spreadsheet: https://docs.google.com/spreadsheets/d/<id>/...")
	flag.Parse()

	args := os.Args
	if sliceIndex(args, "help") >= 0 {
		flag.Usage()
		fmt.Println(`[
 {
   "title": "LoC(Total) by tag",
   "top_left": {
     "x": 900,
     "y": 20
   },
   "size": {
     "height": 380,
     "width": 600
   },
   "x_axis_title": "tag",
   "y_axis_title": "LoC",
   "label_column": "Tag",
   "data_column": "LoC(Total)"
 }
]`)
		return
	}

	title := *titleP
	googleCredentials := *googleCredentialsP
	contentFile := *contentFileP
	chartFile := *chartFileP
	users := *usersP
	highlightColumns := *highlightColumnsP
	sendEmailMessage := *sendEmailMessageP
	spreadsheetId := *spreadsheetIdP

	var userList []string
	if users != "" {
		userList = strings.Split(users, ",")
	}

	table, err := readCsv(contentFile)
	if err != nil {
		log.Fatalf("failed to read csv")
	}
	if len(table) == 0 {
		log.Fatalf("empty csv file")
	}

	ctx := context.Background()

	service, err := api.NewService(ctx, []byte(googleCredentials))
	if err != nil {
		log.Fatalf("failed to start service: %s", err.Error())
	}

	if spreadsheetId != "" {
		err = service.Recreate(spreadsheetId, title)
		if err != nil {
			log.Fatalf("failed to recreate spreadsheet: %s", err.Error())
		}
	} else {
		spreadsheetId, err = service.Create(title)
		if err != nil {
			log.Fatalf("failed to create new spreadsheet: %s", err.Error())
		}
	}

	fmt.Println("https://docs.google.com/spreadsheets/d/" + spreadsheetId)

	for _, user := range userList {
		err = service.Share(spreadsheetId, user)
		if err != nil {
			log.Fatalf("failed to share spreadsheet with user %s: %s", user, err.Error())
		}
	}

	err = service.InsertTable(spreadsheetId, api.Origin(), table)
	if err != nil {
		log.Fatalf("failed to insert table: %s", err.Error())
	}

	highlightColumnsList := strings.Split(highlightColumns, ",")
	if highlightColumns != "" {
		for _, highlightColumn := range highlightColumnsList {
			err = percentileHighlightColumn(service, spreadsheetId, table, highlightColumn)
			if err != nil {
				log.Fatalf("failed to highlight column %s: %s", highlightColumn, err.Error())
			}
		}
	}

	charts, err := api.ReadFromFile(chartFile)
	if err != nil {
		log.Fatalf("failed to read chart config %s", err.Error())
	}
	for _, chart := range charts {
		err = service.AddChart(spreadsheetId, chart)
		if err != nil {
			log.Fatalf("failed to add chart %s: %s", chart.Title, err.Error())
		}
	}

	if sendEmailMessage != "" {
		err = service.SendEmail(users, title, "Document link: https://docs.google.com/spreadsheets/d/"+spreadsheetId+"\n\n"+sendEmailMessage)
		if err != nil {
			log.Fatalf("failed to send email: %s", err.Error())
		}
	}
}

func percentileHighlightColumn(service *api.Service, spreadsheetId string, table [][]string, columnName string) error {
	if len(table) <= 1 {
		return nil
	}

	positiveColor := &api.Color{
		R: 1,
		G: 0.5,
		B: 0.5,
		A: 1,
	}
	neutralColor := &api.Color{
		R: 1,
		G: 1,
		B: 1,
		A: 1,
	}
	negativeColor := &api.Color{
		R: 0.5,
		G: 1,
		B: 0.5,
		A: 1,
	}

	index := sliceIndex(table[0], columnName)
	if index < 0 {
		return errors.New("missing column: " + columnName)
	}
	var negativeValues []float64
	var positiveValues []float64
	for i := 1; i < len(table); i++ {
		value, err := strconv.ParseFloat(table[i][index], 64)
		if err != nil {
			return err
		}
		if value > 0 {
			positiveValues = append(positiveValues, value)
		} else if value < 0 {
			negativeValues = append(negativeValues, value)
		}
	}
	sort.Float64s(negativeValues)
	sort.Float64s(positiveValues)

	percentiles := []float64{
		0.0, 0.2, 0.4, 0.6, 0.8, 1.0,
	}
	if len(negativeValues) > 0 {
		for i := 0; i < len(percentiles)-1; i++ {
			lowerNegativeBoundary := &api.Boundary{Value: negativeValues[int(math.Ceil(float64(len(negativeValues)-1)*percentiles[i]))], Clamp: false}
			upperNegativeBoundary := &api.Boundary{Value: negativeValues[int(math.Ceil(float64(len(negativeValues)-1)*percentiles[i+1]))], Clamp: false}

			currentNegativeColor := api.Lerp(negativeColor, neutralColor, percentiles[i])

			err := service.Highlight(spreadsheetId, api.Origin().Offset(1, index), api.Origin().Offset(len(table)-1, index), lowerNegativeBoundary, upperNegativeBoundary, currentNegativeColor)
			if err != nil {
				return err
			}
		}
	}
	if len(positiveValues) > 0 {
		for i := 0; i < len(percentiles)-1; i++ {
			lowerPositiveBoundary := &api.Boundary{Value: positiveValues[len(positiveValues)-1-int(math.Ceil(float64(len(positiveValues)-1)*percentiles[i]))], Clamp: false}
			upperPositiveBoundary := &api.Boundary{Value: positiveValues[len(positiveValues)-1-int(math.Ceil(float64(len(positiveValues)-1)*percentiles[i+1]))], Clamp: false}

			currentPositiveColor := api.Lerp(positiveColor, neutralColor, percentiles[i])

			err := service.Highlight(spreadsheetId, api.Origin().Offset(1, index), api.Origin().Offset(len(table)-1, index), lowerPositiveBoundary, upperPositiveBoundary, currentPositiveColor)
			if err != nil {
				return err
			}
		}
	}

	return nil
}

func sliceIndex(slice []string, value string) int {
	for index, element := range slice {
		if element == value {
			return index
		}
	}
	return -1
}

func readCsv(file string) ([][]string, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	// read csv values using csv.Reader
	csvReader := csv.NewReader(f)
	data, err := csvReader.ReadAll()
	return data, err
}
