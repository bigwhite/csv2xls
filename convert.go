package main

import (
	"encoding/csv"
	"errors"
	"fmt"
	"os"
	"strconv"
	"unicode"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func getUnitFromCell(cell string) string {
	var idx int
	for _, r := range cell {
		if unicode.IsLetter(r) {
			break
		}
		idx++
	}
	return cell[idx:]
}

func removeUnitFromRecord(records [][]string) error {
	var units []string
	// get Unint from the second row
	for _, cell := range records[1] {
		unit := getUnitFromCell(cell)
		units = append(units, unit)
	}

	// add unit for categories
	categories := records[0]
	for i, cat := range categories {
		if units[i] != "" {
			records[0][i] = cat + "(" + units[i] + ")"
		}
	}

	// remove unit from record
	cells := records
	for i, row := range cells {
		if i != 0 {
			// skip row #1, it is categories
			for j, cell := range row {
				records[i][j] = cell[:len(cell)-len(units[j])]
			}
		}
	}

	return nil
}

func doConvert(input, output string) error {
	f, err := os.Open(input)
	if err != nil {
		return err
	}
	defer f.Close()

	// read all records from csv file
	r := csv.NewReader(f)
	records, err := r.ReadAll()
	if err != nil {
		return err
	}

	if len(records) <= 1 {
		return errors.New("no data or only categories in csv")
	}

	err = removeUnitFromRecord(records)
	if err != nil {
		return err
	}

	fmt.Println(records)

	return nil

	// init cheet of excel file
	ef := excelize.NewFile()
	// set cell value
	//    A  B  C
	// 1  x  x  x
	// 2  x  x  x
	// 3  x  x  x

	for i, record := range records {
		//fmt.Println(record)
		for j, value := range record {
			k := fmt.Sprintf("%c%d", 'A'+j, i+1)
			var s float64
			if s, err = strconv.ParseFloat(value, 32); err == nil {
				err = ef.SetCellFloat("Sheet1", k, s, 2, 32)
				if err != nil {
					return err
				}
			} else {
				err = ef.SetCellValue("Sheet1", k, value)
				if err != nil {
					return err
				}
			}
		}
	}

	// draw a chart

	cell := fmt.Sprintf("A%d", len(records)+1)

	if err = ef.AddChart("Sheet1", cell, `{
          "type": "line",
          "series": [
          {
              "name": "Sheet1!$F$1",
              "categories": "Sheet1!$A$2:$A$7",
              "values": "Sheet1!$F$2:$F$7"
          }
          ],
          "format":
          {
              "x_scale": 1.0,
              "y_scale": 1.0,
              "x_offset": 1,
              "y_offset": 1,
              "print_obj": true,
              "lock_aspect_ratio": false,
              "locked": false
          },
          "legend":
          {
              "position": "top",
              "show_legend_key": false
          },
          "title":
          {
              "name": "CSV chart demo"
          },
          "plotarea":

{
              "show_bubble_size": true,
              "show_cat_name": false,
              "show_leader_lines": false,
              "show_percent": true,
              "show_series_name": false,
              "show_val": false
          },
          "show_blanks_as": "zero",
           "x_axis":
          {
              "reverse_order": false
          },
          "y_axis":
          {
              "maximum": 71360.0,
              "minimum": 71340.0
          }
      }`); err != nil {
		return err
	}

	// save the excel file
	if err = ef.SaveAs("csv.xlsx"); err != nil {
		return err
	}

	return nil
}
