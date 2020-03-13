package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/tealeg/xlsx/v2"
)

func sheet2Map(sheet *xlsx.Sheet) ([]map[string]string, error) {
	if len(sheet.Rows) < 1 {
		return nil, fmt.Errorf("sheet rows error")
	}

	titles := []string{}
	for _, c := range sheet.Rows[0].Cells {
		if len(c.Value) == 0 {
			break
		}
		titles = append(titles, c.Value)
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

func xlsx2Map(xFile *xlsx.File) map[string][]map[string]string {
	responseJSON := map[string][]map[string]string{}
	for _, s := range xFile.Sheets {
		c, err := sheet2Map(s)
		if err != nil {
			continue
		}
		responseJSON[s.Name] = c
	}
	return responseJSON
}

//Convert xlsx to json
func Convert(xFile *xlsx.File) (json.RawMessage, error) {
	data, err := json.Marshal(xlsx2Map(xFile))
	if err != nil {
		return nil, err
	}

	return json.RawMessage(data), nil
}

func main() {
	iName := flag.String("i", "input.xlsx", "input file")
	oName := flag.String("o", "output.json", "output file")
	flag.Parse()

	xFile, err := xlsx.OpenFile(*iName)
	if err != nil {
		log.Fatal(err)
	}

	res, err := Convert(xFile)
	if err != nil {
		log.Fatal(err)
	}

	prettyJSON, err := json.MarshalIndent(res, "", "    ")
	if err != nil {
		log.Fatal("Failed to generate json", err)
	}

	f, err := os.Create(*oName)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer f.Close()
	_, err = f.Write(prettyJSON)
	if err != nil {
		fmt.Println(err)
		f.Close()
		return
	}
}
