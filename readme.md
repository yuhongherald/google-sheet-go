# Google Sheet Go

go1.18+

```
go install github.com/yuhongherald/google-sheet-go
```

go1.17 and older

```
go get github.com/yuhongherald/google-sheet-go
```

## Input options

```
--title=<title>: Title of Google Sheet
--google-sheet-id=<id>: https://docs.google.com/spreadsheets/d/<id>/...
--google-credentials-json=<credentials>: Google credentials JSON string
--cotent-file=<csv filename>: Table contents to be uploaded, in csv format
--highlight-columns=<column-name1,column-name2>: Comma separated column names
--chart-file=<json filename>:list of chart config objects
--users=<users>:comma separated emails
--send-email-message=<message>: Message body for email. Leave blank to skip email sending
```

### Google credentials

Follow the instructions here to generate your credentials JSON file: https://developers.google.com/workspace/guides/create-credentials#service-account

### Chart config

Example
```
[
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
]
```

## Output

Link to Google Sheet
