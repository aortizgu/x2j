package main

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/tealeg/xlsx/v2"
)

var defaultX2JSON = New()

type X2JSON struct {
	toLower bool
	toUpper bool
}

func New() X2JSON {
	return X2JSON{}
}

func (x *X2JSON) ToUpper() *X2JSON {
	x.toUpper = true
	x.toLower = false
	return x
}

func (x *X2JSON) ToLower() *X2JSON {
	x.toUpper = false
	x.toLower = true
	return x
}

func (x *X2JSON) toCase(s string) string {
	if x.toLower {
		return strings.ToLower(s)
	} else if x.toUpper {
		return strings.ToUpper(s)
	}
	return s
}

func (x *X2JSON) sheet2Map(sheet *xlsx.Sheet) ([]map[string]string, error) {
	if len(sheet.Rows) < 1 {
		return nil, fmt.Errorf("sheet rows error")
	}

	titles := []string{}
	for _, c := range sheet.Rows[0].Cells {
		if len(c.Value) == 0 {
			break
		}
		titles = append(titles, x.toCase(c.Value))
	}
	converts := []map[string]string{}
	for _, r := range sheet.Rows[1:] {
		if len(r.Cells[0].Value) == 0 {
			break
		}
		convertMap := map[string]string{}

		for j := 0; j < len(titles); j++ {
			if j >= len(r.Cells) {
				convertMap[titles[j]] = ""
			} else {
				convertMap[titles[j]] = r.Cells[j].Value
			}
		}
		converts = append(converts, convertMap)
	}

	return converts, nil
}

func (x *X2JSON) xlsx2Map(xFile *xlsx.File) map[string][]map[string]string {
	responseJson := map[string][]map[string]string{}
	for _, s := range xFile.Sheets {
		c, err := x.sheet2Map(s)
		if err != nil {
			continue
		}
		responseJson[x.toCase(s.Name)] = c
	}
	return responseJson
}

func (x *X2JSON) Convert(xFile *xlsx.File) (json.RawMessage, error) {
	data, err := json.Marshal(x.xlsx2Map(xFile))
	if err != nil {
		return nil, err
	}

	return json.RawMessage(data), nil
}

func Convert(xFile *xlsx.File) (json.RawMessage, error) {
	return defaultX2JSON.Convert(xFile)
}
