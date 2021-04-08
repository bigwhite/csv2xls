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

func doConvert(input, output string, col int64) error {
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

	//fmt.Println(records)
	return writeToExcelSheet(output, records, col)
}

func writeToExcelSheet(output string, records [][]string, col int64) error {
	var err error

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
			var n int64
			if s, err = strconv.ParseFloat(value, 32); err == nil {
				err = ef.SetCellFloat("Sheet1", k, s, 2, 32)
				if err != nil {
					return err
				}
			} else if n, err = strconv.ParseInt(value, 10, 64); err == nil {
				err = ef.SetCellInt("Sheet1", k, int(n))
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

	cell := fmt.Sprintf("A%d", len(records)+1) // base cell which draw from

	propertiesFmt := `{
		          "type": "line",
		          "series": [
		          {
		              "name": "%s",
		              "categories": "%s",
		              "values": "%s"
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
		              "name": "%s"
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
		          }
		      }`
	/*
	   "name": "Sheet1!$F$1",
	   "categories": "Sheet1!$A$2:$A$7",
	   "values": "Sheet1!$F$2:$F$7"
	*/

	category := records[0][col]
	title := fmt.Sprintf("line chart for %s", category)
	seriesName := fmt.Sprintf(`Sheet1!$%c$%d`, 'A'+col, col)
	seriesCategories := fmt.Sprintf(`Sheet1!$A$2:$A$%d`, len(records))
	seriesValues := fmt.Sprintf(`Sheet1!$%c$2:$%c$%d`, 'A'+col, 'A'+col, len(records))
	//min,max := getMinAndMax(records, col)

	properties := fmt.Sprintf(propertiesFmt, seriesName, seriesCategories, seriesValues, title)

	if err = ef.AddChart("Sheet1", cell, properties); err != nil {
		return err
	}

	// save the excel file
	if err = ef.SaveAs(output); err != nil {
		return err
	}

	return nil
}

/*
func getMinAndMax(records [][]string, col int64) (int, int) {

}
*/
