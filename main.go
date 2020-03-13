package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/tealeg/xlsx/v2"
)

func main() {
	iName := flag.String("i", "input.xlsx", "input file")
	oName := flag.String("o", "output.json", "output file")
	flag.Parse()

	xFile, err := xlsx.OpenFile(*iName)
	if err != nil {
		log.Fatal(err)
	}

	x2j := New()
	res, err := x2j.Convert(xFile)
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
