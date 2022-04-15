package cmd

import (
	"log"
	"os"

	"github.com/sergrom/csv2xls/v2.2/internal/app"
	"github.com/spf13/cobra"
)

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "csv2xls",
	Short: "A command line converter csv into xls",
	Long: `The csv2xls is a command line tool to convert .csv into .xls Excel formats
`,
	Run: func(cmd *cobra.Command, args []string) {
		var csvFileName, xlsFileName string
		var err error
		if csvFileName, err = cmd.Flags().GetString("csv-file-name"); err != nil {
			log.Fatal("Please specify csv-file-name parameter")
		}

		if xlsFileName, err = cmd.Flags().GetString("xls-file-name"); err != nil {
			log.Fatal("Please specify xls-file-name parameter")
		}

		csvDelimiter := ";"
		delim, err := cmd.Flags().GetString("csv-delimiter")
		if err != nil {
			log.Fatal(err.Error())
		}
		if len(delim) > 0 {
			csvDelimiter = delim
		}

		var title, subject, creator, keywords, description, lastModifiedBy string

		if title, err = cmd.Flags().GetString("title"); err != nil {
			log.Fatal(err.Error())
		}
		if subject, err = cmd.Flags().GetString("subject"); err != nil {
			log.Fatal(err.Error())
		}
		if creator, err = cmd.Flags().GetString("creator"); err != nil {
			log.Fatal(err.Error())
		}
		if keywords, err = cmd.Flags().GetString("keywords"); err != nil {
			log.Fatal(err.Error())
		}
		if description, err = cmd.Flags().GetString("description"); err != nil {
			log.Fatalf(err.Error())
		}
		if lastModifiedBy, err = cmd.Flags().GetString("last-modified-by"); err != nil {
			log.Fatal(err.Error())
		}

		converter, err := app.NewCsv2XlsConverter(csvFileName, xlsFileName, csvDelimiter)
		if err != nil {
			log.Fatal(err.Error())
		}

		err = converter.
			WithTitle(title).
			WithSubject(subject).
			WithDescription(description).
			WithKeywords(keywords).
			WithCreator(creator).
			WithLastModifiedBy(lastModifiedBy).
			Convert()

		if err != nil {
			log.Fatal(err.Error())
		}
	},
}

// Execute ...
func Execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	// Mandatory parameter
	rootCmd.Flags().String("csv-file-name", "", `The input csv file you want to convert`)
	_ = rootCmd.MarkFlagRequired("csv-file-name")

	// Mandatory parameter
	rootCmd.Flags().String("xls-file-name", "", `The output xls file name that will be created`)
	_ = rootCmd.MarkFlagRequired("xls-file-name")

	// Optional parameters:
	rootCmd.Flags().String("csv-delimiter", "", `Optional. The delimiter that used in csv file. Default value is semicolon - ";"`)
	rootCmd.Flags().String("title", "", `Optional. The Title property of xls file`)
	rootCmd.Flags().String("subject", "", `Optional. The Subject property of xls file`)
	rootCmd.Flags().String("creator", "", `Optional. The Creator property of xls file`)
	rootCmd.Flags().String("keywords", "", `Optional. The Keywords property of xls file`)
	rootCmd.Flags().String("description", "", `Optional. The Description property of xls file`)
	rootCmd.Flags().String("last-modified-by", "", `Optional. The LastModifiedBy property of xls file`)
}
