package main

import (
	"flag"
	"os"
	"fmt"
	"github.com/tealeg/xlsx"
	"strings"
	"strconv"
	"github.com/olekukonko/tablewriter"
)


func min(a, b int64) int64 {
	if a > b {
		return b
	}
	return a
}

func main() {
	sheet := flag.String("s", "1", "sheet name or number")
	rn := flag.String("r", "1,10", "row numbers")
	cn := flag.String("c", "1,10", "cell numbers,  you must give number instead of letters")
	flag.Parse()

	args := os.Args[1:]

	fname := args[len(args)-1]

	if 0 == len(fname) {
		fmt.Println("file name is missing")
		os.Exit(-1)
	}

	rns := strings.Split(*rn, ",")
	sr, err := strconv.ParseInt(rns[0], 10, 64)
	if err != nil {
		fmt.Println("rows start number is not a number")
		os.Exit(-1)
	}

	er, err := strconv.ParseInt(rns[1], 10, 64)
	if err != nil {
		fmt.Println("rows end number is not a number")
		os.Exit(-1)
	}

	cns := strings.Split(*cn, ",")
	sc, err := strconv.ParseInt(cns[0], 10, 64)
	if err != nil {
		fmt.Println("cell start number is not a number")
		os.Exit(-1)
	}

	ec, err := strconv.ParseInt(cns[1], 10, 64)
	if err != nil {
		fmt.Println("cell end number is not a number")
		os.Exit(-1)
	}

	fp, err := xlsx.OpenFile(fname)
	if err != nil {
		fmt.Println("file not found")
		os.Exit(-1)
	}
	var sh *xlsx.Sheet

	// they give us sheet number
 	if pn, err := strconv.ParseInt(*sheet, 10, 64); err == nil {
		if  0 < pn {
			pn -= 1
		}

		sh = fp.Sheets[pn]

	} else {
		// they give us sheet's name (how they know sheets name?, who cares, they know somehow)
		var ok bool
		sh, ok = fp.Sheet[*sheet]
		if !ok {
			fmt.Println("sheet not found")
			os.Exit(-1)
		}
	}

	if 0 < sr {
		sr -= 1
	}
	if 0 < sc {
		sc -= 1
	}

	table := tablewriter.NewWriter(os.Stdout)

	for _, row := range sh.Rows[sr:min(int64(len(sh.Rows)), int64(er - 1))] {
		data := make([]string, 0)
		for _, cell := range row.Cells[sc:min(ec - 1, int64(len(row.Cells)))] {
			if str, err := cell.FormattedValue(); err == nil {
				data = append(data, strings.TrimSpace(str[0:min(int64(15), int64(len(str)))]))
			}
		}
		table.Append(data)
	}

	table.SetRowLine(true)
	table.SetRowSeparator("-")
	table.Render()
}
