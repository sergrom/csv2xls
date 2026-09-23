package cmd

import (
	"fmt"

	"github.com/sergrom/csv2xls/v3/internal/app"
	"github.com/spf13/cobra"
)

// NewRootCommand creates an independent command with its own flag values.
func NewRootCommand() *cobra.Command {
	var input, output, delimiter string
	var title, subject, creator, keywords, description, lastModifiedBy string
	command := &cobra.Command{
		Use:           "csv2xls",
		Short:         "Convert CSV files to XLS workbooks",
		Args:          cobra.NoArgs,
		SilenceUsage:  true,
		SilenceErrors: true,
		RunE: func(cmd *cobra.Command, args []string) error {
			if input == "" || output == "" {
				return fmt.Errorf("csv-file-name and xls-file-name must not be empty")
			}
			// Preserve the previous behavior for an explicitly empty delimiter.
			if delimiter == "" {
				delimiter = ";"
			}
			converter, err := app.NewCsv2XlsConverter(input, output, delimiter)
			if err != nil {
				return err
			}
			return converter.WithTitle(title).WithSubject(subject).WithCreator(creator).
				WithKeywords(keywords).WithDescription(description).
				WithLastModifiedBy(lastModifiedBy).Convert()
		},
	}
	flags := command.Flags()
	flags.StringVar(&input, "csv-file-name", "", "Input CSV file (required)")
	flags.StringVar(&output, "xls-file-name", "", "Output XLS file (required)")
	_ = command.MarkFlagRequired("csv-file-name")
	_ = command.MarkFlagRequired("xls-file-name")
	flags.StringVar(&delimiter, "csv-delimiter", ";", "CSV field delimiter")
	flags.StringVar(&title, "title", "", "Document title")
	flags.StringVar(&subject, "subject", "", "Document subject")
	flags.StringVar(&creator, "creator", "", "Document creator")
	flags.StringVar(&keywords, "keywords", "", "Document keywords")
	flags.StringVar(&description, "description", "", "Document description")
	flags.StringVar(&lastModifiedBy, "last-modified-by", "", "Last modified by")
	return command
}

// Execute runs the command and returns errors to the caller.
func Execute() error { return NewRootCommand().Execute() }
