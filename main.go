package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
)

var (
	input  = flag.String("i", "", "the name of csv file")
	output = flag.String("o", "", "the name of xls file")
)

const (
	xlsExt = ".xlsx"
)

// csv2xls -i xxx.csv // -> xxx.xlsx
// csv2xls -i xxx.csv -o yyy.xlsx
func main() {
	flag.Usage = Usage
	flag.Parse()

	if *input == "" {
		fmt.Fprintln(os.Stderr, "no input csv file specified. Use -i arg to specify input csv file")
		Usage()
		os.Exit(1)
	}

	if *output == "" {
		ext := filepath.Ext(*input)
		*output = (*input)[:len(*input)-len(ext)] + xlsExt
	}

	err := doConvert(*input, *output)
	if err != nil {
		panic(err)
	}
	fmt.Printf("convert [%s] to [%s] ok\n", *input, *output)
}

// Usage reimplements flag.Usage
func Usage() {
	progname := os.Args[0]
	fmt.Fprintf(os.Stderr, "Usage of %s:\n", progname)
	flag.PrintDefaults()
	fmt.Fprintf(os.Stderr, `
Examples:
        %s -i xxx.csv
        %s -i xxx.csv -o yyy.xlsx
`, progname, progname)
}
