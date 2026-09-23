/*
Copyright © 2022 Sergey Romanenko <serega170587@gmail.com>
*/
package main

import (
	"fmt"
	"os"

	"github.com/sergrom/csv2xls/v3/cmd"
)

func main() {
	if err := cmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, "Error:", err)
		os.Exit(1)
	}
}
