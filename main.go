package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"path"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gorilla/mux"
	uuid "github.com/satori/go.uuid"
)

// SheetName - name of sheet
const SheetName = "Sheet1"

func main() {
	port := os.Getenv("PORT")
	if port == "" {
		port = "5000"
	}

	r := mux.NewRouter()
	r.HandleFunc("/", XLSXFormatter)

	serverPath := fmt.Sprintf("localhost:%s", port)
	log.Printf("Start server on http://%s\n", serverPath)
	if err := http.ListenAndServe(serverPath, r); err != nil {
		log.Fatal("ListenAndServe: ", err)
	}
}

// XLSXFormatter - create and return xlsx file
func XLSXFormatter(w http.ResponseWriter, r *http.Request) {
	err := r.ParseForm()
	if err != nil {
		log.Println("Bad Request", err)
	}

	var columns []string
	for _, column := range r.Form["columns[]"] {
		columns = append(columns, column)
	}

	type record []string
	var records []record

	for key, value := range r.Form {
		if strings.HasPrefix(key, "data[") {
			rec := make(record, 0, len(columns))
			rec = append(rec, value...)
			records = append(records, rec)
		}
	}

	xlsx := excelize.NewFile()

	for index, column := range columns {
		xlsx.SetCellValue(SheetName, toCharStr(index+1)+"1", column)
	}

	for index, record := range records {
		for recordIndex, recordValue := range record {
			xlsx.SetCellValue(
				SheetName,
				toCharStr(recordIndex+1)+strconv.FormatInt(int64(index+2), 10),
				recordValue,
			)
		}
	}

	fileName := uuid.NewV4().String() + ".xlsx"
	fp := path.Join("files", fileName)
	xlsx.SaveAs(fp)

	w.Header().Set("Content-Disposition", "attachment; filename=Result.xlsx")
	file, err := ioutil.ReadFile(fp)
	w.Write(file)
}

func toCharStr(i int) string {
	return string('A' - 1 + i)
}
