package main

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"reflect"
	"sort"
	"strconv"
	"strings"
)

func main() {
	inputFile, err := excelize.OpenFile("input.xlsx")
	if err != nil {
		println(err.Error())
		return
	}

	outFile := excelize.NewFile()
	// Create a new sheet.
	index := outFile.NewSheet("Лист1")

	outFile.SetCellValue("Лист1", "B1", "Насел. Пункт")
	outFile.SetCellValue("Лист1", "C1", "Улица")
	outFile.SetCellValue("Лист1", "D1", "Номер дома")
	outFile.SetCellValue("Лист1", "E1", "Номер квартиры")
	outFile.SetCellValue("Лист1", "G1", "ФИО жильцов")
	outFile.SetCellValue("Лист1", "H1", "собственника")
	outFile.SetCellValue("Лист1", "i1", "признак прописнного")

	var liv []string
	var sob []string
	var adr []string
	var nstr int = 2
	rows, err := inputFile.GetRows("Лист1")
	rows[1][1]
	for _, row := range rows {
		liv = strings.Split(row[6], "\n")
		sob = strings.Split(row[4], "\n")
		adr = strings.Split(row[1], ",")
		sort.Strings(sob)
		sort.Strings(liv)
		for i := 0; i <len(sob); i++ {
			if len(sob[i])<2{break}
			outFile.SetCellValue("Лист1", "B"+strconv.Itoa(nstr), adr[0])
			outFile.SetCellValue("Лист1", "C"+strconv.Itoa(nstr), adr[1])
			outFile.SetCellValue("Лист1", "D"+strconv.Itoa(nstr), adr[2])
			outFile.SetCellValue("Лист1", "E"+strconv.Itoa(nstr), adr[3])
			outFile.SetCellValue("Лист1", "G"+strconv.Itoa(nstr), sob[i])
			outFile.SetCellValue("Лист1", "H"+strconv.Itoa(nstr), "1")
			if itemExists(liv, sob[i]) {
				outFile.SetCellValue("Лист1", "I"+strconv.Itoa(nstr), "1")
			}
			nstr++
		}

		for i := 0; i < len(liv); i++ {
			if len(liv[i])<2{break}
			if itemExists(sob, liv[i])  {

			} else {

			outFile.SetCellValue("Лист1", "B"+strconv.Itoa(nstr), adr[0])
			outFile.SetCellValue("Лист1", "C"+strconv.Itoa(nstr), adr[1])
			outFile.SetCellValue("Лист1", "D"+strconv.Itoa(nstr), adr[2])
			outFile.SetCellValue("Лист1", "E"+strconv.Itoa(nstr), adr[3])
			outFile.SetCellValue("Лист1", "G"+strconv.Itoa(nstr), liv[i])
			outFile.SetCellValue("Лист1", "I"+strconv.Itoa(nstr), "1")
			if itemExists(sob, liv[i])  {
				outFile.SetCellValue("Лист1", "H"+strconv.Itoa(nstr), "1")
			}
			nstr++
		}
		}
	}
	outFile.SetActiveSheet(index)
	if err := outFile.SaveAs("out.xlsx"); err != nil {
		println(err.Error())
	}
}


func itemExists(slice interface{}, item interface{}) bool {
	s := reflect.ValueOf(slice)

	if s.Kind() != reflect.Slice {
		panic("Invalid data-type")
	}

	for i := 0; i < s.Len(); i++ {
		if s.Index(i).Interface() == item {
			return true
		}
	}

	return false
}
