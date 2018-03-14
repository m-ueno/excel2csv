package main

import (
	"encoding/csv"
	"flag"
	"log"
	"os"
	"time"

	"github.com/tealeg/xlsx"
)

func sheetToFileName(sheet *xlsx.Sheet) string {
	return sheet.Name + ".csv"
}

func sheetToCSVFile(sheet *xlsx.Sheet) error {
	csvFile, err := os.Create(sheetToFileName(sheet))
	defer csvFile.Close()

	if err != nil {
		return err
	}

	w := csv.NewWriter(csvFile)
	defer w.Flush()

	for _, row := range sheet.Rows {
		var vals []string

		for _, cell := range row.Cells {
			str, err := cell.FormattedValue()
			if err != nil {
				return err
			}
			vals = append(vals, str)
		}

		if err := w.Write(vals); err != nil {
			return nil
		}
	}

	log.Printf("debug: sheet [%s] sucessfully converted to [%s]", sheet.Name, sheetToFileName(sheet))

	return nil
}

func xlsx2CSVFiles(filename string) error {
	xlFile, err := xlsx.OpenFile(filename)
	if err != nil {
		return err
	}

	for _, sheet := range xlFile.Sheets {
		if err := sheetToCSVFile(sheet); err != nil {
			return err
		}
	}

	return nil
}

func runInterval(filename string, interval time.Duration) error {
	var fInfo os.FileInfo
	var err error

	if fInfo, err = os.Stat(filename); err != nil {
		return err
	}
	lastUpdate := fInfo.ModTime().UnixNano()

	log.Printf("debug: watching [%s]\n", filename)

	t := time.NewTicker(interval)
	defer t.Stop()

	for {
		select {
		case <-t.C:
			if fInfo, err = os.Stat(filename); err != nil {
				return err
			}
			modTime := fInfo.ModTime().UnixNano()

			if modTime > lastUpdate {
				if err = xlsx2CSVFiles(filename); err != nil {
					return err
				}
				lastUpdate = modTime
			}
		}
	}
}

func main() {
	/*
		input: filename
		output: csv files
	*/
	var input string
	var watch bool
	flag.StringVar(&input, "input", "", "parse `xlsx` file")
	flag.BoolVar(&watch, "watch", true, "watch and convert xlsx file periodically")
	flag.Parse()

	if input == "" {
		flag.PrintDefaults()
		os.Exit(1)
	}

	// 1st run
	if err := xlsx2CSVFiles(input); err != nil {
		log.Fatal(err)
	}

	if watch {
		if err := runInterval(input, 5*time.Second); err != nil {
			log.Fatal(err)
		}
	}
}
