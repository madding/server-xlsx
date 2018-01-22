package main

import (
	"encoding/json"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"path"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
	uuid "github.com/satori/go.uuid"
)

// SheetName - name of sheet
const SheetName = "Sheet1"

func main() {
	port := os.Getenv("PORT")
	if port == "" {
		port = "5000"
	}

	http.HandleFunc("/", XLSXFormatter)

	serverPath := fmt.Sprintf("localhost:%s", port)
	log.Printf("Start server on http://%s\n", serverPath)
	if err := http.ListenAndServe(serverPath, nil); err != nil {
		log.Fatal("ListenAndServe: ", err)
	}
}

// XLSXFormatter - create and return xlsx file
func XLSXFormatter(w http.ResponseWriter, r *http.Request) {
	params := ParseRequestBody(r.Body)

	xlsx := excelize.NewFile()

	for index, column := range params.Fields {
		xlsx.SetCellValue(SheetName, toCharStr(index+1)+"1", column.Name)
		style, err := xlsx.NewStyle(`{"alignment":{"horizontal":"center"},"font":{"bold":true}}`)
		if err != nil {
			log.Println("Error create style", err)
		}
		xlsx.SetCellStyle(SheetName, toCharStr(index+1)+"1", toCharStr(index+1)+"1", style)
	}

	for rowIndex, position := range params.Data {
		for colIndex, column := range params.Fields {
			var recordValue interface{}
			for columnName, recValue := range position {
				if columnName == column.ColumnName {
					recordValue = recValue
				}
			}

			if recordValue != nil {
				var styleParams Style
				cellNumber := toCharStr(colIndex+1) + strconv.FormatInt(int64(rowIndex+2), 10)

				styleParams.Alignment.Vertical = "center"
				if column.TableAlign != "" {
					styleParams.Alignment.Horizontal = column.TableAlign
				}

				switch column.Type {
				case "Int":
					style, _ := xlsx.NewStyle(styleParams.ToJSON())
					xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)

					numberValue, OK := recordValue.(float64)
					if OK {
						xlsx.SetCellValue(SheetName, cellNumber, int64(numberValue))
					} else {
						xlsx.SetCellValue(SheetName, cellNumber, recordValue)
					}
				case "Select":
					selectValue, OK := recordValue.(float64)
					style, _ := xlsx.NewStyle(styleParams.ToJSON())
					xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)

					if !OK {
						log.Printf("Select calue not int column %v, option %v \n", column.Options, recordValue)
						xlsx.SetCellValue(SheetName, cellNumber, recordValue)
						break
					}

					optionValue, OK := column.Options[strconv.FormatInt(int64(selectValue), 10)]
					if !OK {
						log.Printf("Option not found column %v, option %v \n", column.Options, selectValue)
						xlsx.SetCellValue(SheetName, cellNumber, selectValue)
						break
					}
					xlsx.SetCellValue(SheetName, cellNumber, optionValue)
				case "MoneyWithCurrency":
					stringValue, OK := recordValue.(string)
					if !OK {
						style, _ := xlsx.NewStyle(styleParams.ToJSON())
						xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)
						xlsx.SetCellValue(SheetName, cellNumber, recordValue)
						break
					}
					splittedStr := strings.Split(stringValue, " ")
					floatValue, err := strconv.ParseFloat(splittedStr[1], 10)
					if err == nil {
						styleParams.CustomNumberFormat = fmt.Sprintf("#,##0.00\\ [$%s]", splittedStr[0])
						style, err := xlsx.NewStyle(styleParams.ToJSON())
						if err != nil {
							log.Println("Can't set custom number format", styleParams.ToJSON())
						}

						xlsx.SetCellValue(SheetName, cellNumber, floatValue)
						xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)
					} else {
						log.Printf("Number with currency not float column: %v, value: %v", column, recordValue)
						style, _ := xlsx.NewStyle(styleParams.ToJSON())
						xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)
						xlsx.SetCellValue(SheetName, cellNumber, recordValue)
					}
				case "Date":
					stringValue, OK := recordValue.(string)

					if OK {
						dateValue, err := time.Parse("2006-01-02", stringValue)
						if err != nil {
							dates := strings.Split(stringValue, "\n")
							var res []string
							for _, date := range dates {
								parsedDate, err := time.Parse("2006-01-02", date)
								if err != nil {
									log.Printf("Error parse date %v format 2006-01-02\n", date)
								} else {
									res = append(res, parsedDate.Format("02.01.2006"))
								}
							}

							styleParams.Alignment.WrapText = true
							style, _ := xlsx.NewStyle(styleParams.ToJSON())
							xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)
							xlsx.SetCellValue(SheetName, cellNumber, strings.Join(res, "\n"))
							break
						}

						styleParams.CustomNumberFormat = "[$-380A]dd.mm.yyyy"
						style, _ := xlsx.NewStyle(styleParams.ToJSON())
						if err != nil {
							log.Println(err)
						}
						xlsx.SetCellValue(SheetName, cellNumber, dateValue)
						xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)
					} else {
						style, _ := xlsx.NewStyle(styleParams.ToJSON())
						xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)
						xlsx.SetCellValue(SheetName, cellNumber, recordValue)
					}
				default:
					style, _ := xlsx.NewStyle(styleParams.ToJSON())
					xlsx.SetCellStyle(SheetName, cellNumber, cellNumber, style)
					stringValue, OK := recordValue.(string)
					if OK {
						xlsx.SetCellValue(SheetName, cellNumber, stringValue)
					} else {
						xlsx.SetCellValue(SheetName, cellNumber, recordValue)
					}
				}
			} else {
				xlsx.SetCellValue(
					SheetName,
					toCharStr(colIndex+1)+strconv.FormatInt(int64(rowIndex+2), 10),
					"Колонка не найдена в файле",
				)
			}
		}
	}

	fileName := uuid.NewV4().String() + ".xlsx"
	fp := path.Join("files", fileName)
	xlsx.SaveAs(fp)

	file, err := ioutil.ReadFile(fp)
	if err != nil {
		log.Println("Can't create xlsx file")
	}

	w.Header().Set("Content-Disposition", "attachment; filename=Result.xlsx")
	w.Write(file)
}

func toCharStr(i int) string {
	return string('A' - 1 + i)
}

// ParseRequestBody parse POST request body, for getting form params
func ParseRequestBody(ioBody io.ReadCloser) (jsonBody ParamsBody) {
	body, err := ioutil.ReadAll(ioBody)
	if err != nil {
		log.Println("Error parse request body ", err)
	}

	json.Unmarshal(body, &jsonBody)
	return
}
