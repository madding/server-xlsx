# Simple server take params and return xslx file
- fields - column names
- data - data for xlsx

## Start server
```
  go run *.go
```

## Post request
```
  curl -X POST -d @test/test.json localhost:5000
```
XLSX file generated via JSON file in directory files/ and sended back to client
