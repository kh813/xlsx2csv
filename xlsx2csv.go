package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
)

const (
	VERSION  = "0.1.0"
)

func main() {
	var (
		filename string
		sheet    int
		tsv  bool
		help bool
		win  bool
	)
	flag.StringVar(&filename, "f", "", "File to read")
	flag.IntVar(&sheet, "s", 0, "Sheet number")
	flag.BoolVar(&help, "h", false, "Show help")
	flag.BoolVar(&tsv, "t", false, "TSV output")
	flag.BoolVar(&win, "w", false, "CRLF for Windows")
	flag.Parse()

	if help == true || len(os.Args) <= 1 {
		showhelp()
	}

	// Adjust sheet number
	if sheet >= 1 {
		sheet--
	}

	// Open Excel file
	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Get all the rows in the sheet
	rows, err := f.GetRows(f.GetSheetName(sheet))
	if err != nil {
		fmt.Println(err)
		return
	}

	if tsv == true {
		for _, row := range rows {
			for _, colCell := range row {
				fmt.Printf("%s\t", colCell)
			}
			fmt.Println()
		}
	} else {
		// CSV output
		w := csv.NewWriter(os.Stdout)
		if win == true {
			w.UseCRLF = true
		}
		for _, row := range rows {
			if err := w.Write(row); err != nil {
				fmt.Println("error writing record to csv:", err)
			}
		}
		w.Flush()
	}
}

func showhelp() {
	fmt.Println(`Covert Excel(xlsx) file to CSV

Usage:
    xlsx2csv -f <Excel file> [other flags]

Flags : 
    -f Excel_file.xlsx       Specify the file to import 
    -s "(1 or more)"         Sheet number
    -t  (default: false)     TSV output flag  
    -w  (default: false)     Windows CRLF flag, only works with -c flag
    -h                       Show help
	
Example: 
    Show "my_file.xlsx" in CSV format
        ./xlsx2csv -f my_file.xlsx
    Show 2nd sheet of "my_file.xlsx" in CSV format and redirect to a file
        ./xlsx2csv -f my_file.xlsx -s 2 > new_file.csv
    Show "my_file.xlsx" in TSV format
        ./xlsx2csv -f my_file.xlsx -t 
`)
	os.Exit(0)
}

