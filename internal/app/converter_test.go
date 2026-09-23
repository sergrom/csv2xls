package app

import (
	"bytes"
	"encoding/binary"
	"errors"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"unicode/utf16"
)

func encoded(s string) []byte {
	var b bytes.Buffer
	_ = binary.Write(&b, binary.LittleEndian, utf16.Encode([]rune(s)))
	return b.Bytes()
}

func TestDelimiter(t *testing.T) {
	for _, delimiter := range []string{"", "ab", "\x00", "\n", "\r", "\"", "\xff", "�"} {
		if _, err := NewCsv2XlsConverter("", "", delimiter); err == nil {
			t.Errorf("accepted %q", delimiter)
		}
	}
	for _, delimiter := range []string{";", ",", "\t", "|", "§"} {
		if _, err := NewCsv2XlsConverter("", "", delimiter); err != nil {
			t.Errorf("rejected %q: %v", delimiter, err)
		}
	}
}

func TestConvert(t *testing.T) {
	dir := t.TempDir()
	input := filepath.Join(dir, "input.csv")
	output := filepath.Join(dir, "output.xls")
	// Different record lengths and a bare quote preserve the original CSV settings.
	if err := os.WriteFile(input, []byte("City;Code\nParis;001\nBerlin\nHello\"world;🐧\n"), 0600); err != nil {
		t.Fatal(err)
	}
	c, err := NewCsv2XlsConverter(input, output, ";")
	if err != nil {
		t.Fatal(err)
	}
	c.WithTitle("Report").WithSubject("Cities").WithCreator("Author").WithKeywords("Travel").WithDescription("Example").WithLastModifiedBy("Editor")
	if err := c.Convert(); err != nil {
		t.Fatal(err)
	}
	data, err := os.ReadFile(output)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.HasPrefix(data, []byte{0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1}) {
		t.Fatal("missing XLS container")
	}
	for _, value := range []string{"worksheet", "Paris", "001", "Berlin", "Hello\"world", "🐧", "Report", "Cities", "Author", "Travel", "Example", "Editor"} {
		if !bytes.Contains(data, encoded(value)) {
			t.Errorf("missing %q", value)
		}
	}
}

func TestSheetBoundaries(t *testing.T) {
	for _, count := range []int{0, rowsPerSheet, rowsPerSheet + 1} {
		t.Run(strconv.Itoa(count), func(t *testing.T) {
			dir := t.TempDir()
			input := filepath.Join(dir, "input.csv")
			output := filepath.Join(dir, "output.xls")
			if err := os.WriteFile(input, []byte(strings.Repeat("value\n", count)), 0600); err != nil {
				t.Fatal(err)
			}
			c, _ := NewCsv2XlsConverter(input, output, ";")
			if err := c.Convert(); err != nil {
				t.Fatal(err)
			}
			data, err := os.ReadFile(output)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Contains(data, encoded("worksheet")) {
				t.Fatal("missing first sheet")
			}
			if bytes.Contains(data, encoded("worksheet1")) != (count > rowsPerSheet) {
				t.Fatal("incorrect split boundary")
			}
		})
	}
}

func TestConversionErrors(t *testing.T) {
	dir := t.TempDir()
	input := filepath.Join(dir, "input.csv")
	output := filepath.Join(dir, "output.xls")
	c, _ := NewCsv2XlsConverter(input, output, ";")
	if err := c.Convert(); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("missing input: %v", err)
	}
	for _, content := range []string{strings.Repeat("a;", 256) + "a\n", strings.Repeat("x", 32768) + "\n"} {
		if err := os.WriteFile(input, []byte(content), 0600); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(output, []byte("existing output"), 0600); err != nil {
			t.Fatal(err)
		}
		if err := c.Convert(); err == nil || !strings.Contains(err.Error(), "record 1") {
			t.Fatalf("invalid record: %v", err)
		}
		data, _ := os.ReadFile(output)
		if string(data) != "existing output" {
			t.Fatal("invalid input overwrote output")
		}
	}
	if err := os.WriteFile(input, []byte("valid\n"), 0600); err != nil {
		t.Fatal(err)
	}
	c.xlsFileName = dir
	if err := c.Convert(); err == nil {
		t.Fatal("output error not returned")
	}
}
