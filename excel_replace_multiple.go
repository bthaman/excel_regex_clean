package main

import (
	"dirs"
	"fmt"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/naoina/toml"
)

type tomlConfig struct {
	Title  string
	Search struct {
		Wb_regex            string
		Sheet_regex         string
		Cell_regex          string
		Searchreplace_regex [][]string
		Replace_blanks		string
		Value_if_blank      string
	}
}

func main() {
	f, err := os.Open("excel_replace_multiple.toml")
	if err != nil {
		panic(err)
	}
	defer f.Close()
	var config tomlConfig
	if err := toml.NewDecoder(f).Decode(&config); err != nil {
		panic(err)
	}
	wb_regex := config.Search.Wb_regex
	sheet_regex := config.Search.Sheet_regex
	cell_regex := config.Search.Cell_regex
	searchreplace_regex := config.Search.Searchreplace_regex
	replace_blanks := config.Search.Replace_blanks
	value_if_blank := config.Search.Value_if_blank

	files, _ := dirs.Walk_Dir_Info(".")
	for i, f := range files {
		matched, _ := regexp.Match(wb_regex, []byte(f))
		if matched {
			fmt.Println(i, f)
			clean_data(f, sheet_regex, cell_regex, searchreplace_regex, replace_blanks, value_if_blank)
		}
	}
}

func clean_data(fn string, sheet_regex string, cell_regex string, searchreplace_regex [][]string, replace_blanks string, value_if_blank string) {
	xlsx, err := excelize.OpenFile(fn)
	if err != nil {
		fmt.Println(err)
		return
	}
	var cell_address string
	// loop through worksheets and apply regex to matched sheets/cells
	for index, name := range xlsx.GetSheetMap() {
		matched_sheet, _ := regexp.Match(sheet_regex, []byte(name))
		if matched_sheet {
			fmt.Println(index, name)
			rows, _ := xlsx.GetRows(name)
			for r, row := range rows {
				for c, colCell := range row {
					if len(colCell) >= 0 {
						cell_address, _ = CoordinatesToCellName(c+1, r+1)
						match_cell, _ := regexp.Match(cell_regex, []byte(cell_address))
						if match_cell {
							// replace cell values using array of regex/replace values
							for i := 0; i < len(searchreplace_regex); i++ {
								match_val, _ := regexp.Match(searchreplace_regex[i][0], []byte(colCell))
								if match_val {
									m := regexp.MustCompile(searchreplace_regex[i][0])
									val_new := m.ReplaceAllString(strings.Trim(colCell, " "), searchreplace_regex[i][1])
									// re-assign colCell so that successive ReplaceAllStrings can be applied to the last value
									colCell = val_new
									// attempt to convert val_new to a float
									if flt, err := strconv.ParseFloat(val_new, 32); err != nil {
										// if error converting to float, set cell to text value
										xlsx.SetCellValue(name, cell_address, val_new)
									} else {
										// if val_new successfully converted to float, assign the float to cell
										xlsx.SetCellFloat(name, cell_address, flt, 4, 32)
									}
								}
							}
						}
					} 
					if replace_blanks == "TRUE" {
						if len(colCell) == 0 {
							cell_address, _ = CoordinatesToCellName(c+1, r+1)
							// set cell to value_if_blank
							if flt, err := strconv.ParseFloat(value_if_blank, 32); err != nil {
								// if error converting to float, set cell to text value
								xlsx.SetCellValue(name, cell_address, value_if_blank)
							} else {
								// if val_new successfully converted to float, assign the float to cell
								xlsx.SetCellFloat(name, cell_address, flt, 4, 32)
							}	
						}
					}
				}
			}
		}
	}
	if err := xlsx.Save(); err != nil {
		println(err.Error())
	}
}

func ColumnNumberToName(num int) (string, error) {
	if num < 1 {
		return "", fmt.Errorf("incorrect column number %d", num)
	}
	if num > 16384 {
		return "", fmt.Errorf("column number exceeds maximum limit")
	}
	var col string
	for num > 0 {
		col = string(rune((num-1)%26+65)) + col
		num = (num - 1) / 26
	}
	return col, nil
}

func CoordinatesToCellName(col, row int) (string, error) {
	if col < 1 || row < 1 {
		return "", fmt.Errorf("invalid cell coordinates [%d, %d]", col, row)
	}
	colname, err := ColumnNumberToName(col)
	return colname + strconv.Itoa(row), err
}
