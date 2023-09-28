package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func main() {
	// Open the Excel file
	file, err := excelize.OpenFile("sample.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	// Get all the sheet names
	sheets := file.GetSheetMap()

	// Loop through each sheet
	for _, sheet := range sheets {
		rows, err := file.GetRows(sheet)
		if err != nil {
			fmt.Println(err)
			return
		}

		// Loop through each row
		for _, row := range rows {
			// Loop through each cell
			for _, cell := range row {
				style, err := file.GetCellStyle(sheet, cell)
				if err != nil {
					fmt.Println(err)
					return
				}

				// Check if the font color is red (ColorIndex 3)
				if style.Font.Color.Indexed == 3 {
					fmt.Println("Found a cell with red font color:", cell)
				}
			}
		}
	}
}
