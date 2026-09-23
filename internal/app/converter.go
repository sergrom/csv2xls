package app

import (
	"bufio"
	"bytes"
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"strconv"
	"unicode/utf8"

	"github.com/sergrom/xls-writer"
)

// Keep the existing converter's sheet boundaries for compatibility.
const rowsPerSheet = 65535

// Csv2XlsConverter ...
type Csv2XlsConverter struct {
	csvFileName    string
	xlsFileName    string
	csvDelimiter   rune
	title          string
	subject        string
	creator        string
	keywords       string
	description    string
	lastModifiedBy string
}

// NewCsv2XlsConverter configures CSV input and XLS output paths.
func NewCsv2XlsConverter(csvFileName, xlsFileName, csvDelimiter string) (*Csv2XlsConverter, error) {
	delimiter, size := utf8.DecodeRuneInString(csvDelimiter)
	if size != len(csvDelimiter) || size == 0 || delimiter == utf8.RuneError || delimiter == 0 || delimiter == '"' || delimiter == '\r' || delimiter == '\n' {
		return nil, fmt.Errorf("csv delimiter must be one valid character other than a quote, newline, or NUL")
	}
	return &Csv2XlsConverter{csvFileName: csvFileName, xlsFileName: xlsFileName, csvDelimiter: delimiter}, nil
}

// Convert reads CSV records and saves a workbook using xls-writer.
func (c *Csv2XlsConverter) Convert() error {
	f, err := os.Open(c.csvFileName)
	if err != nil {
		return fmt.Errorf("open CSV: %w", err)
	}
	defer f.Close()

	input := bufio.NewReader(f)
	// Strip only an initial UTF-8 BOM; BOM characters inside fields are data.
	if prefix, _ := input.Peek(3); bytes.Equal(prefix, []byte{0xEF, 0xBB, 0xBF}) {
		_, _ = input.Discard(3)
	}
	reader := csv.NewReader(input)
	reader.Comma = c.csvDelimiter
	reader.FieldsPerRecord = -1
	reader.LazyQuotes = true
	reader.ReuseRecord = true // AppendRow copies each record.

	w := xlswriter.New()
	w.SetProperties(xlswriter.Properties{
		Title: c.title, Subject: c.subject, Creator: c.creator,
		Keywords: c.keywords, Description: c.description, LastModifiedBy: c.lastModifiedBy,
	})
	sheet, err := w.AddSheet("worksheet")
	if err != nil {
		return err
	}
	for row := 0; ; row++ {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("read CSV record %d: %w", row+1, err)
		}
		if row > 0 && row%rowsPerSheet == 0 {
			sheet, err = w.AddSheet("worksheet" + strconv.Itoa(row/rowsPerSheet))
			if err != nil {
				return fmt.Errorf("create sheet for CSV record %d: %w", row+1, err)
			}
		}
		if err := sheet.AppendRow(record); err != nil {
			return fmt.Errorf("append CSV record %d: %w", row+1, err)
		}
	}
	if err := w.Save(c.xlsFileName); err != nil {
		return fmt.Errorf("save XLS: %w", err)
	}
	return nil
}

// WithTitle ...
func (c *Csv2XlsConverter) WithTitle(title string) *Csv2XlsConverter {
	c.title = title
	return c
}

// WithSubject ...
func (c *Csv2XlsConverter) WithSubject(subject string) *Csv2XlsConverter {
	c.subject = subject
	return c
}

// WithCreator ...
func (c *Csv2XlsConverter) WithCreator(creator string) *Csv2XlsConverter {
	c.creator = creator
	return c
}

// WithKeywords ...
func (c *Csv2XlsConverter) WithKeywords(keywords string) *Csv2XlsConverter {
	c.keywords = keywords
	return c
}

// WithDescription ...
func (c *Csv2XlsConverter) WithDescription(description string) *Csv2XlsConverter {
	c.description = description
	return c
}

// WithLastModifiedBy ...
func (c *Csv2XlsConverter) WithLastModifiedBy(lastModifiedBy string) *Csv2XlsConverter {
	c.lastModifiedBy = lastModifiedBy
	return c
}
