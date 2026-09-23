package cmd

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestCommandErrors(t *testing.T) {
	dir := t.TempDir()
	input := filepath.Join(dir, "input.csv")
	output := filepath.Join(dir, "output.xls")
	if err := os.WriteFile(input, []byte("City;Code\nParis;001\n"), 0600); err != nil {
		t.Fatal(err)
	}
	base := []string{"--csv-file-name", input, "--xls-file-name", output}
	cases := []struct {
		name    string
		args    []string
		message string
	}{
		{"required", nil, "required flag"},
		{"positional", append(append([]string{}, base...), "unexpected"), "unknown command"},
		{"empty input", []string{"--csv-file-name=", "--xls-file-name", output}, "must not be empty"},
		{"delimiter", append(append([]string{}, base...), "--csv-delimiter=ab"), "csv delimiter"},
		{"missing input", []string{"--csv-file-name", filepath.Join(dir, "missing.csv"), "--xls-file-name", output}, "open CSV"},
		{"bad output", []string{"--csv-file-name", input, "--xls-file-name", dir}, "save XLS"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			command := NewRootCommand()
			var buf bytes.Buffer
			command.SetOut(&buf)
			command.SetErr(&buf)
			command.SetArgs(tc.args)
			err := command.Execute()
			if err == nil || !strings.Contains(err.Error(), tc.message) {
				t.Fatalf("got %v, want %q", err, tc.message)
			}
			if buf.Len() != 0 {
				t.Fatalf("command printed an error instead of returning it: %s", buf.String())
			}
		})
	}
}

func TestCommandInstances(t *testing.T) {
	dir := t.TempDir()
	input := filepath.Join(dir, "input.csv")
	output := filepath.Join(dir, "output.xls")
	if err := os.WriteFile(input, []byte("City;Code\nParis;001\n"), 0600); err != nil {
		t.Fatal(err)
	}
	for _, extra := range [][]string{nil, {"--csv-delimiter="}} {
		command := NewRootCommand()
		command.SetArgs(append([]string{"--csv-file-name", input, "--xls-file-name", output}, extra...))
		if err := command.Execute(); err != nil {
			t.Fatal(err)
		}
	}
	// Required-flag state must not leak from a previous instance.
	command := NewRootCommand()
	command.SetArgs(nil)
	if err := command.Execute(); err == nil {
		t.Fatal("new command inherited flags")
	}
}

func TestHelp(t *testing.T) {
	command := NewRootCommand()
	var buf bytes.Buffer
	command.SetOut(&buf)
	command.SetArgs([]string{"--help"})
	if err := command.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(buf.String(), `(default ";")`) {
		t.Fatalf("missing delimiter default in help: %s", buf.String())
	}
}
