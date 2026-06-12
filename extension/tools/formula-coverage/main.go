// Command formula-coverage generates FORMULAS.md: PhpSpreadsheet's calculation
// functions (phpspreadsheet-functions.txt, extracted from
// Calculation/FunctionArray.php) probed against excelize's formula engine,
// which easy-excel delegates getCalculatedValue() to.
//
//	go run ./tools/formula-coverage > ../FORMULAS.md
//
// A function counts as supported when excelize parses and dispatches it (any
// error other than its "not support ... function" rejection: argument-count
// or domain errors prove the function exists).
package main

import (
	_ "embed"
	"fmt"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
)

//go:embed phpspreadsheet-functions.txt
var phpFunctions string

func main() {
	f := excelize.NewFile()
	defer f.Close()
	_ = f.SetCellValue("Sheet1", "A1", 1)
	_ = f.SetCellValue("Sheet1", "A2", 2)

	names := strings.Fields(phpFunctions)
	sort.Strings(names)

	var supported, missing []string
	for _, name := range names {
		if probe(f, name) {
			supported = append(supported, name)
		} else {
			missing = append(missing, name)
		}
	}

	fmt.Println("# Formula function coverage")
	fmt.Println()
	fmt.Println("`getCalculatedValue()` / `toArray(calculateFormulas: true)` delegate to")
	fmt.Println("[excelize](https://github.com/qax-os/excelize)'s formula engine. This table")
	fmt.Println("probes every function PhpSpreadsheet's calculation engine registers")
	fmt.Println("(extracted from `Calculation/FunctionArray.php`) against that engine —")
	fmt.Println("regenerate with `go run ./tools/formula-coverage > ../FORMULAS.md` from")
	fmt.Println("`extension/`.")
	fmt.Println()
	fmt.Printf("**%d of %d** PhpSpreadsheet functions are available (%d missing).\n",
		len(supported), len(names), len(missing))
	fmt.Println()
	fmt.Println("Semantics of individual functions may still differ in edge cases")
	fmt.Println("(see COMPAT.md); cached results in loaded files are always preferred.")
	fmt.Println()
	fmt.Println("## Not available via excelize")
	fmt.Println()
	for _, name := range missing {
		fmt.Printf("- `%s`\n", name)
	}
	fmt.Println()
	fmt.Println("## Available")
	fmt.Println()
	fmt.Println("| | | | |")
	fmt.Println("|---|---|---|---|")
	for i := 0; i < len(supported); i += 4 {
		row := make([]string, 4)
		for j := 0; j < 4 && i+j < len(supported); j++ {
			row[j] = "`" + supported[i+j] + "`"
		}
		fmt.Printf("| %s |\n", strings.Join(row, " | "))
	}
}

// probe reports whether excelize's engine knows the function: unknown names
// fail with exactly "not support <NAME> function", known names succeed or
// fail on argument validation. A few engine functions panic on the dummy
// arguments — a panic still proves the function was dispatched.
func probe(f *excelize.File, name string) (known bool) {
	if err := f.SetCellFormula("Sheet1", "C1", name+"(A1,A2)"); err != nil {
		return false
	}
	defer func() {
		if recover() != nil {
			known = true
		}
	}()
	_, err := f.CalcCellValue("Sheet1", "C1")
	if err == nil {
		return true
	}
	return err.Error() != "not support "+name+" function"
}
