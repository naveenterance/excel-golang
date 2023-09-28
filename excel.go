package main

import (
	"fmt"
	"log"

	"github.com/tealeg/xlsx"
)

func main() {
	// Open the Excel file
	file, err := xlsx.OpenFile("exx.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Read data from the sheet
	for _, sheet := range file.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				fmt.Printf("%s\t", cell.String())
			}
			fmt.Println()
		}
	}
}
