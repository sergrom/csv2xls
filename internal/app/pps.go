package app

import (
	"bytes"
	"time"
)

// pps ...
type pps struct {
	No         int
	Name       string
	PpsType    uint8
	PrevPps    uint32
	NextPps    uint32
	DirPps     uint32
	Data       string
	Size       uint32
	StartBlock uint32
}

// getPpsWk ...
func (pps *pps) getPpsWk() string {
	buf := new(bytes.Buffer)
	putVar(buf, []byte(padRight(pps.Name, "\x00", 64)))

	putVar(buf,
		int16(len(pps.Name)+2),
		pps.PpsType,
		int8(0x00),
		pps.PrevPps,
		pps.NextPps,
		pps.DirPps,
		[]byte("\x00\x09\x02\x00"),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\xc0\x00\x00\x00"),
		[]byte("\x00\x00\x00\x46"),
		[]byte("\x00\x00\x00\x00"),
		[]byte(localDateToOLE(time.Now().Unix())),
		[]byte(localDateToOLE(time.Now().Unix())),
		pps.StartBlock,
		pps.Size,
		uint32(0),
	)

	return buf.String()
}
