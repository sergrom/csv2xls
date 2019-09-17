package main

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"math"
	"unicode/utf16"
	"unicode/utf8"
)

func PadRight(str, pad string, lenght int) string {
	for {
		str += pad
		if len(str) > lenght {
			return str[0:lenght]
		}
	}
}

func putVar(w io.Writer, args ...interface{}) {
	for _, i := range args {
		err := binary.Write(w, binary.LittleEndian, i)
		if err != nil {
			fmt.Println(err)
			panic("Cannot write int16")
		}
	}
}

func localDateToOLE(timestamp int64) string {
	var factor int64 = 4294967296
	var days int64 = 134774
	bigDate :=  days * 24 * 3600 + timestamp + 10800
	bigDate *= 10000000

	highPart := int64(float64(bigDate) / float64(factor))
	// lower 4 bytes
	lowPart := int(math.Floor(((float64(bigDate) / float64(factor)) - float64(highPart)) * float64(factor)))

	buf := new(bytes.Buffer)
	var hex int
	for i := 0; i < 4; i++ {
		hex = lowPart % 256
		putVar(buf, uint8(hex))
		lowPart = int(math.Floor(float64(lowPart) / 256))
	}
	for i := 0; i < 4; i++ {
		hex = int(highPart) % 256
		putVar(buf, uint8(hex))
		highPart = int64(math.Floor(float64(highPart) / 0x100))
	}

	return buf.String();
}

//Utility function to transform ASCII text to Unicode.
func ascToUcs(ascii string) string {
	buf := new(bytes.Buffer)
	for i:=0; i<len(ascii); i++ {
		putVar(buf, ascii[i], []byte("\x00"))
	}

	return buf.String()
}

// Converts a UTF-8 string into BIFF8 Unicode string data (8-bit string length)
func UTF8toBIFF8UnicodeShort(value string) string {
	buf := new(bytes.Buffer)
	ln := utf8.RuneCountInString(value)
	utf16str := utf16.Encode([]rune(value))
	putVar(buf, uint8(ln), uint8(0x0001), utf16str)

	return buf.String()
}

// Converts a UTF-8 string into BIFF8 Unicode string data (16-bit string length)
func UTF8toBIFF8UnicodeLong(value string) string {
	buf := new(bytes.Buffer)
	ln := utf8.RuneCountInString(value)
	utf16str := utf16.Encode([]rune(value))
	putVar(buf, uint16(ln), uint8(0x0001), utf16str)

	return buf.String()
}

// Max returns the larger of x or y.
func Max(x, y int) int {
	if x < y {
		return y
	}
	return x
}
func MaxUInt16(x, y uint16) uint16 {
	if x < y {
		return y
	}
	return x
}

func MinUInt16(x, y uint16) uint16 {
	if x > y {
		return y
	}
	return x
}

func substr(slice []byte, start, length int) []byte  {
	return slice[start:start+length]
}