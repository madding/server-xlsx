package main

import (
	"encoding/json"
)

// ParamsBody - params file
type ParamsBody struct {
	Fields []struct {
		ColumnName string            `json:"key"`
		Type       string            `json:"type"`
		Name       string            `json:"name"`
		TableAlign string            `json:"tableAlign,omitempty"`
		Options    map[string]string `json:"options,omitempty"`
	} `json:"fields"`
	Data []struct {
		Name                     string `json:"Name"`
		CompanyID                int    `json:"CompanyID"`
		DocType                  int    `json:"DocType"`
		Status                   int    `json:"Status"`
		DocDate                  string `json:"DocDate"`
		CacheLinesContainerName  string `json:"CacheLinesContainerName"`
		Amount                   string `json:"Amount"`
		Overdue                  int    `json:"Overdue"`
		ShipmentETS              string `json:"ShipmentETS"`
		Comments                 string `json:"Comments"`
		LinesTotalOriginalAmount string `json:"LinesTotalOriginalAmount"`
		ID                       int    `json:"ID"`
	} `json:"data"`
}

// Style - style model
type Style struct {
	Alignment struct {
		Horizontal string `json:"horizontal,omitempty"`
		Vertical   string `json:"vertical,omitempty"`
		WrapText   bool   `json:"wrap_text,omitempty"`
	} `json:"alignment"`
	CustomNumberFormat string `json:"custom_number_format,omitempty"`
}

// ToJSON - export to JSON
func (style Style) ToJSON() string {
	res, err := json.Marshal(style)
	if err != nil {
		return ""
	}

	return string(res)
}
