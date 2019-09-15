package main

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"strings"
)

type Rgb struct {
	index        int
	red          uint8
	green        uint8
	blue         uint8
	transparent  uint8
}

var wbPallete = []Rgb{
	Rgb{0x08, 0x00, 0x00, 0x00, 0x00},
	Rgb{0x09, 0xff, 0xff, 0xff, 0x00},
	Rgb{0x0A, 0xff, 0x00, 0x00, 0x00},
	Rgb{0x0B, 0x00, 0xff, 0x00, 0x00},
	Rgb{0x0C, 0x00, 0x00, 0xff, 0x00},
	Rgb{0x0D, 0xff, 0xff, 0x00, 0x00},
	Rgb{0x0E, 0xff, 0x00, 0xff, 0x00},
	Rgb{0x0F, 0x00, 0xff, 0xff, 0x00},
	Rgb{0x10, 0x80, 0x00, 0x00, 0x00},
	Rgb{0x11, 0x00, 0x80, 0x00, 0x00},
	Rgb{0x12, 0x00, 0x00, 0x80, 0x00},
	Rgb{0x13, 0x80, 0x80, 0x00, 0x00},
	Rgb{0x14, 0x80, 0x00, 0x80, 0x00},
	Rgb{0x15, 0x00, 0x80, 0x80, 0x00},
	Rgb{0x16, 0xc0, 0xc0, 0xc0, 0x00},
	Rgb{0x17, 0x80, 0x80, 0x80, 0x00},
	Rgb{0x18, 0x99, 0x99, 0xff, 0x00},
	Rgb{0x19, 0x99, 0x33, 0x66, 0x00},
	Rgb{0x1A, 0xff, 0xff, 0xcc, 0x00},
	Rgb{0x1B, 0xcc, 0xff, 0xff, 0x00},
	Rgb{0x1C, 0x66, 0x00, 0x66, 0x00},
	Rgb{0x1D, 0xff, 0x80, 0x80, 0x00},
	Rgb{0x1E, 0x00, 0x66, 0xcc, 0x00},
	Rgb{0x1F, 0xcc, 0xcc, 0xff, 0x00},
	Rgb{0x20, 0x00, 0x00, 0x80, 0x00},
	Rgb{0x21, 0xff, 0x00, 0xff, 0x00},
	Rgb{0x22, 0xff, 0xff, 0x00, 0x00},
	Rgb{0x23, 0x00, 0xff, 0xff, 0x00},
	Rgb{0x24, 0x80, 0x00, 0x80, 0x00},
	Rgb{0x25, 0x80, 0x00, 0x00, 0x00},
	Rgb{0x26, 0x00, 0x80, 0x80, 0x00},
	Rgb{0x27, 0x00, 0x00, 0xff, 0x00},
	Rgb{0x28, 0x00, 0xcc, 0xff, 0x00},
	Rgb{0x29, 0xcc, 0xff, 0xff, 0x00},
	Rgb{0x2A, 0xcc, 0xff, 0xcc, 0x00},
	Rgb{0x2B, 0xff, 0xff, 0x99, 0x00},
	Rgb{0x2C, 0x99, 0xcc, 0xff, 0x00},
	Rgb{0x2D, 0xff, 0x99, 0xcc, 0x00},
	Rgb{0x2E, 0xcc, 0x99, 0xff, 0x00},
	Rgb{0x2F, 0xff, 0xcc, 0x99, 0x00},
	Rgb{0x30, 0x33, 0x66, 0xff, 0x00},
	Rgb{0x31, 0x33, 0xcc, 0xcc, 0x00},
	Rgb{0x32, 0x99, 0xcc, 0x00, 0x00},
	Rgb{0x33, 0xff, 0xcc, 0x00, 0x00},
	Rgb{0x34, 0xff, 0x99, 0x00, 0x00},
	Rgb{0x35, 0xff, 0x66, 0x00, 0x00},
	Rgb{0x36, 0x66, 0x66, 0x99, 0x00},
	Rgb{0x37, 0x96, 0x96, 0x96, 0x00},
	Rgb{0x38, 0x00, 0x33, 0x66, 0x00},
	Rgb{0x39, 0x33, 0x99, 0x66, 0x00},
	Rgb{0x3A, 0x00, 0x33, 0x00, 0x00},
	Rgb{0x3B, 0x33, 0x33, 0x00, 0x00},
	Rgb{0x3C, 0x99, 0x33, 0x00, 0x00},
	Rgb{0x3D, 0x99, 0x33, 0x66, 0x00},
	Rgb{0x3E, 0x33, 0x33, 0x99, 0x00},
	Rgb{0x3F, 0x33, 0x33, 0x33, 0x00},
}
//var wbPallete = map[int][]uint8{
//	0x08 : {0x00, 0x00, 0x00, 0x00},
//	0x09 : {0xff, 0xff, 0xff, 0x00},
//	0x0A : {0xff, 0x00, 0x00, 0x00},
//	0x0B : {0x00, 0xff, 0x00, 0x00},
//	0x0C : {0x00, 0x00, 0xff, 0x00},
//	0x0D : {0xff, 0xff, 0x00, 0x00},
//	0x0E : {0xff, 0x00, 0xff, 0x00},
//	0x0F : {0x00, 0xff, 0xff, 0x00},
//	0x10 : {0x80, 0x00, 0x00, 0x00},
//	0x11 : {0x00, 0x80, 0x00, 0x00},
//	0x12 : {0x00, 0x00, 0x80, 0x00},
//	0x13 : {0x80, 0x80, 0x00, 0x00},
//	0x14 : {0x80, 0x00, 0x80, 0x00},
//	0x15 : {0x00, 0x80, 0x80, 0x00},
//	0x16 : {0xc0, 0xc0, 0xc0, 0x00},
//	0x17 : {0x80, 0x80, 0x80, 0x00},
//	0x18 : {0x99, 0x99, 0xff, 0x00},
//	0x19 : {0x99, 0x33, 0x66, 0x00},
//	0x1A : {0xff, 0xff, 0xcc, 0x00},
//	0x1B : {0xcc, 0xff, 0xff, 0x00},
//	0x1C : {0x66, 0x00, 0x66, 0x00},
//	0x1D : {0xff, 0x80, 0x80, 0x00},
//	0x1E : {0x00, 0x66, 0xcc, 0x00},
//	0x1F : {0xcc, 0xcc, 0xff, 0x00},
//	0x20 : {0x00, 0x00, 0x80, 0x00},
//	0x21 : {0xff, 0x00, 0xff, 0x00},
//	0x22 : {0xff, 0xff, 0x00, 0x00},
//	0x23 : {0x00, 0xff, 0xff, 0x00},
//	0x24 : {0x80, 0x00, 0x80, 0x00},
//	0x25 : {0x80, 0x00, 0x00, 0x00},
//	0x26 : {0x00, 0x80, 0x80, 0x00},
//	0x27 : {0x00, 0x00, 0xff, 0x00},
//	0x28 : {0x00, 0xcc, 0xff, 0x00},
//	0x29 : {0xcc, 0xff, 0xff, 0x00},
//	0x2A : {0xcc, 0xff, 0xcc, 0x00},
//	0x2B : {0xff, 0xff, 0x99, 0x00},
//	0x2C : {0x99, 0xcc, 0xff, 0x00},
//	0x2D : {0xff, 0x99, 0xcc, 0x00},
//	0x2E : {0xcc, 0x99, 0xff, 0x00},
//	0x2F : {0xff, 0xcc, 0x99, 0x00},
//	0x30 : {0x33, 0x66, 0xff, 0x00},
//	0x31 : {0x33, 0xcc, 0xcc, 0x00},
//	0x32 : {0x99, 0xcc, 0x00, 0x00},
//	0x33 : {0xff, 0xcc, 0x00, 0x00},
//	0x34 : {0xff, 0x99, 0x00, 0x00},
//	0x35 : {0xff, 0x66, 0x00, 0x00},
//	0x36 : {0x66, 0x66, 0x99, 0x00},
//	0x37 : {0x96, 0x96, 0x96, 0x00},
//	0x38 : {0x00, 0x33, 0x66, 0x00},
//	0x39 : {0x33, 0x99, 0x66, 0x00},
//	0x3A : {0x00, 0x33, 0x00, 0x00},
//	0x3B : {0x33, 0x33, 0x00, 0x00},
//	0x3C : {0x99, 0x33, 0x00, 0x00},
//	0x3D : {0x99, 0x33, 0x66, 0x00},
//	0x3E : {0x33, 0x33, 0x99, 0x00},
//	0x3F : {0x33, 0x33, 0x33, 0x00},
//}

type Workbook struct {
	WorksheetSizes      []int
	WorksheetNames      []string
	stringCollection    *StringCollection
}

func (wb *Workbook) getWorksheetSizesData() string {
	buf := new(bytes.Buffer)

	// Calculate the number of selected worksheet tabs and call the finalization
	// methods for each worksheet
	totalWorksheets := len(wb.WorksheetSizes);

	// Add part 1 of the Workbook globals, what goes before the SHEET records
	wb.storeBof(buf)
	wb.writeCodepage(buf)
	wb.writeWindow1(buf)

	wb.writeDateMode(buf)
	wb.writeAllFonts(buf)
	wb.writeAllNumberFormats(buf)
	wb.writeAllXfs(buf)
	wb.writeAllStyles(buf)
	wb.writePalette(buf)

	// Prepare part 3 of the workbook global stream, what goes after the SHEET records
	part3Buf := new(bytes.Buffer)

	wb.writeRecalcId(part3Buf)
	
	wb.writeSupbookInternal(part3Buf, totalWorksheets);
	/* TODO: store external SUPBOOK records and XCT and CRN records
	   in case of external references for BIFF8 */
	wb.writeExternalsheetBiff8(part3Buf, totalWorksheets);
	wb.writeAllDefinedNamesBiff8(part3Buf);
	wb.writeMsoDrawingGroup(part3Buf);
	wb.writeSharedStringsTable(part3Buf);

	wb.writeEof(part3Buf);

	// Add part 2 of the Workbook globals, the SHEET records
	worksheetOffsets := wb.calcSheetOffsets(buf.Len() + part3Buf.Len(), totalWorksheets)
	for i := 0; i < totalWorksheets; i++ {
		wb.writeBoundSheet(buf, wb.WorksheetNames[i], worksheetOffsets[i])
	}

	// Add part 3 of the Workbook globals
	buf.Write(part3Buf.Bytes())

	return buf.String()
}

func (wb *Workbook) storeBof(buffer *bytes.Buffer) {
	var wbType uint16 = 0x0005

	var record uint16 = 0x0809 // Record identifier    (BIFF5-BIFF8)
	var length uint16 = 0x0010

	var build uint16 = 0x0DBB //    Excel 97
	var year uint16 = 0x07CC //    Excel 97

	var version uint16 = 0x0600 //    BIFF8

	putVar(buffer, record, length)
	putVar(buffer, version, wbType, build, year)

	// by inspection of real files, MS Office Excel 2007 writes the following
	putVar(buffer, uint32(0x000100D1), uint32(0x00000406))
}

func (wb *Workbook) writeCodepage(buffer *bytes.Buffer) {
	var record uint16 = 0x0042 // Record identifier
	var length uint16 = 0x0002 // Number of bytes to follow
	var cv uint16 = 0x04B0 // The code page

	putVar(buffer, record, length, cv)
}

func (wb *Workbook) writeWindow1(buffer *bytes.Buffer) {
	var record uint16 = 0x003D // Record identifier
	var length uint16 = 0x0012 // Number of bytes to follow

	var xWn uint16 = 0x0000 // Horizontal position of window
	var yWn uint16 = 0x0000 // Vertical position of window
	var dxWn uint16 = 0x25BC // Width of window
	var dyWn uint16 = 0x1572 // Height of window

	var grbit uint16 = 0x0038 // Option flags

	// not supported by PhpSpreadsheet, so there is only one selected sheet, the active
	var ctabsel uint16 = 1 // Number of workbook tabs selected

	var wTabRatio uint16 = 0x0258 // Tab to scrollbar ratio

	// not supported by PhpSpreadsheet, set to 0
	var itabFirst uint16 = 0 // 1st displayed worksheet
	var itabCur uint16 = 0  // Active worksheet

	putVar(buffer, record, length)
	putVar(buffer, xWn, yWn, dxWn, dyWn, grbit, itabCur, itabFirst, ctabsel, wTabRatio)
}

func (wb *Workbook) writeDateMode(buffer *bytes.Buffer) {
	var record uint16 = 0x0022; // Record identifier
	var length uint16 = 0x0002; // Bytes to follow

	var f1904 uint16 = 0 // Flag for 1904 date system

	putVar(buffer, record, length, f1904)
}

func (wb *Workbook) writeAllFonts(buffer *bytes.Buffer) {
	var icv uint16 = 8 // Index to color palette
	var sss uint16 = 0

	var bFamily uint8 = 0 // Font family
	var bCharSet uint8 = 0x00 // Character set

	var record uint16 = 0x31 // Record identifier
	var reserved uint8 = 0x00 // Reserved
	var grbit uint16 = 0x00 // Font attributes

	dataBuf := new(bytes.Buffer)

	var fontSize uint16 = 11
	putVar(dataBuf,
		fontSize*20,
		grbit,
		icv,                 // Colour
		uint16(0x190),       // Font weight (0x190=400=normal)
		sss,                 // Superscript/Subscript
		uint8(0x00),         // Underline
		bFamily,
		bCharSet,
		reserved,
		[]byte(UTF8toBIFF8UnicodeShort("Calibri")),
	)

	putVar(buffer, record, uint16(dataBuf.Len()))
	buffer.Write(dataBuf.Bytes())
}

func (wb *Workbook) writeAllNumberFormats(buffer *bytes.Buffer) {
	// empty
}

func (wb *Workbook) writeAllXfs(buffer *bytes.Buffer) {
	var record uint16 = 0x00E0 // Record identifier
	var length uint16 = 0x0014 // Number of bytes to follow

	for i := 0; i < 15; i++ {
		putVar(buffer, record, length)
		putVar(buffer, uint16(0), uint16(0), uint16(0xFFF5), uint8(32))
		putVar(buffer, uint8(0), uint8(0), uint8(0xC0))
		putVar(buffer, uint32(0), uint32(0), uint16(1033))
	}

	putVar(buffer, record, length)
	putVar(buffer, uint16(0), uint16(0), uint16(1), uint8(32))
	putVar(buffer, uint8(0), uint8(0), uint8(0xC0))
	putVar(buffer, uint32(0), uint32(0), uint16(1033))
}

func (wb *Workbook) writeAllStyles(buffer *bytes.Buffer) {
	var record uint16 = 0x0293 // Record identifier
	var length uint16 = 0x0004 // Bytes to follow

	var ixfe uint16 = 0x8000 // Index to cell style XF
	var BuiltIn uint8 = 0x00 // Built-in style
	var iLevel uint8 = 0xff // Outline style level

	putVar(buffer, record, length)
	putVar(buffer, ixfe, BuiltIn, iLevel)
}

func (wb *Workbook) writePalette(buffer *bytes.Buffer) {
	var record uint16 = 0x0092 // Record identifier
	length := 2 + 4 * len(wbPallete) // Number of bytes to follow
	ccv := len(wbPallete) // Number of RGB values to follow

	putVar(buffer, record, uint16(length), uint16(ccv))

	// Pack the RGB data
	for _, color := range wbPallete {
		putVar(buffer, color.red, color.green, color.blue, color.transparent)
	}
}

func (wb *Workbook) writeRecalcId(buffer *bytes.Buffer) {
	var record uint16 = 0x01C1 // Record identifier
	var length uint16 = 8 // Number of bytes to follow

	putVar(buffer, record, length)

	// by inspection of real Excel files, MS Office Excel 2007 writes this
	putVar(buffer, uint32(0x000001C1), uint32(0x00001E667))
}

func (wb *Workbook) writeSupbookInternal(buffer *bytes.Buffer, totalWorksheets int) {
	var record uint16 = 0x01AE // Record identifier
	var length uint16 = 0x0004 // Bytes to follow

	putVar(buffer, record, length)
	putVar(buffer, uint16(totalWorksheets),  uint16(0x0401))
}

func (wb *Workbook) writeExternalsheetBiff8(buffer *bytes.Buffer, totalWorksheets int) {
	if totalWorksheets > 255 {
		panic("Too many worksheets")
	}

	tmpBuf := new(bytes.Buffer)

	cWorksheets := uint16(totalWorksheets)

	var record uint16 = 0x0017 // Record identifier
	var length uint16 = 2 + 6 * cWorksheets // Number of bytes to follow

	putVar(tmpBuf, record, length)
	putVar(tmpBuf, cWorksheets)

	var i uint16
	for i=0 ; i < cWorksheets; i++ {
		putVar(tmpBuf, uint16(0x00),i ,i)
	}

	wb.writeData(buffer, tmpBuf)
}

func (wb *Workbook) writeAllDefinedNamesBiff8(buffer *bytes.Buffer) {
	// empty
}

func (wb *Workbook) writeMsoDrawingGroup(buffer *bytes.Buffer) {
	// empty
}

func (wb *Workbook) writeData(bufferTo *bytes.Buffer, bufferFrom *bytes.Buffer) {
	if bufferFrom.Len() -4 > 8224 {
		wb.addContinue(bufferTo, bufferFrom)
	}

	bufferTo.Write(bufferFrom.Bytes())
}

func (wb *Workbook) addContinue(bufferTo *bytes.Buffer, bufferFrom *bytes.Buffer) {
	var limit uint16 = 8224
	var record uint16 = 0x003C // Record identifier

	putVar(bufferTo, substr(bufferFrom.Bytes(), 0 ,2), limit, substr(bufferFrom.Bytes(), 4, 8224))

	bufFromLength := bufferFrom.Len()

	var i int
	for i = int(limit+4); i < (bufFromLength - int(limit));  i += int(limit) {
		putVar(bufferTo, record, limit)
		putVar(bufferTo, substr(bufferFrom.Bytes(), i, int(limit)))
	}

	// Retrieve the last chunk of data
	putVar(bufferTo, record, uint16(bufferFrom.Len() - i), bufferFrom.Bytes()[i:])
}

func (wb *Workbook) writeSharedStringsTable(buffer *bytes.Buffer) {
	// maximum size of record data (excluding record header)
	continueLimit := 8224

	// initialize array of record data blocks
	recordDatas :=  make([]string, 0)

	// start SST record data block with total number of strings, total number of unique strings
	var recordData strings.Builder

	//var data strings.Builder

	//data.WriteString(workbook.getWorksheetSizesData())
	buf := new(bytes.Buffer)
	putVar(buf, uint32(wb.stringCollection.stringTotal), uint32(wb.stringCollection.stringUnique))
	recordData.Write(buf.Bytes())

	for _,str := range wb.stringCollection.stringList {

		var length uint16
		var encoding uint8
		reader := bytes.NewReader([]byte(str))
		err1 := binary.Read(reader, binary.LittleEndian, &length)
		if err1 != nil {
			fmt.Println("binary.Read failed:", err1)
		}
		err2 := binary.Read(reader, binary.LittleEndian, &encoding)
		if err2 != nil {
			fmt.Println("binary.Read failed:", err2)
		}

		finished := false
		for finished == false {
			// normally, there will be only one cycle, but if string cannot immediately be written as is
			// there will be need for more than one cylcle, if string longer than one record data block, there
			// may be need for even more cycles

			if recordData.Len() + len(str) <= continueLimit {

				recordData.WriteString(str)

				if recordData.Len() + len(str) == continueLimit {
					// we close the record data block, and initialize a new one
					recordDatas = append(recordDatas, recordData.String())
					recordData.Reset()
				}

				// we are finished writing this string
				finished = true
			} else {
				// special treatment writing the string (or remainder of the string)
				// If the string is very long it may need to be written in more than one CONTINUE record.

				// check how many bytes more there is room for in the current record
				spaceRemaining := continueLimit - recordData.Len()

				// minimum space needed
				// uncompressed: 2 byte string length length field + 1 byte option flags + 2 byte character
				// compressed:   2 byte string length length field + 1 byte option flags + 1 byte character
				minSpaceNeeded := 5

				// We have two cases
				// 1. space remaining is less than minimum space needed
				//        here we must waste the space remaining and move to next record data block
				// 2. space remaining is greater than or equal to minimum space needed
				//        here we write as much as we can in the current block, then move to next record data block

				// 1. space remaining is less than minimum space needed
				if spaceRemaining < minSpaceNeeded {

					// we close the block, store the block data
					recordDatas = append(recordDatas, recordData.String())

					// and start new record data block where we start writing the string
					recordData.Reset()

					// 2. space remaining is greater than or equal to minimum space needed
				} else {
					// initialize effective remaining space, for Unicode strings this may need to be reduced by 1, see below
					effectiveSpaceRemaining := spaceRemaining

					// for uncompressed strings, sometimes effective space remaining is reduced by 1
					if encoding == 1 && (len(str) - spaceRemaining) % 2 == 1 {
						effectiveSpaceRemaining--
					}

					// one block fininshed, store the block data
					recordData.WriteString(str[0:effectiveSpaceRemaining])

					str = str[effectiveSpaceRemaining:] // for next cycle in while loop
					recordDatas = append(recordDatas, recordData.String())

					// start new record data block with the repeated option flags
					recordData.Reset()
					recordData.Write([]byte("\x01"))// putVar(recordData, encoding)
				}
			}
		}
	}

	// Store the last record data block unless it is empty
	// if there was no need for any continue records, this will be the for SST record data block itself
	if recordData.Len() > 0 {
		recordDatas = append(recordDatas, recordData.String())
	}

	// combine into one chunk with all the blocks SST, CONTINUE,...
	for i, rData := range recordDatas {
		// first block should have the SST record header, remaing should have CONTINUE header
		var record uint16 = 0x003C
		if i == 0 {
			record = 0x00FC
		}

		tmpBuf := new(bytes.Buffer)
		putVar(tmpBuf, record, uint16(len(rData)), []byte(rData))

		wb.writeData(buffer, tmpBuf)
	}
}

func (wb *Workbook) writeEof(buffer *bytes.Buffer) {
	var record uint16 = 0x000A // Record identifier
	var length uint16 = 0x0000 // Number of bytes to follow

	putVar(buffer, record, length)
}

func (wb *Workbook) calcSheetOffsets(dataSize int, totalWorksheets int) []uint32 {
	worksheetOffsets := make([]uint32, 0)
	boundSheetLength := 10 // fixed length for a BOUNDSHEET record

	// size of Workbook globals part 1 + 3
	offset := dataSize

	// add size of Workbook globals part 2, the length of the SHEET records
	for _, sheetTitle := range wb.WorksheetNames {
		offset += boundSheetLength + len(UTF8toBIFF8UnicodeShort(sheetTitle))
	}

	// add the sizes of each of the Sheet substreams, respectively
	for i := 0; i < totalWorksheets; i++ {
		worksheetOffsets = append(worksheetOffsets, uint32(offset))
		offset += wb.WorksheetSizes[i]
	}

	return worksheetOffsets
}

func (wb *Workbook) writeBoundSheet(buffer *bytes.Buffer, sheetName string, offset uint32) {
	var record uint16 = 0x0085 // Record identifier
	var ss uint8 = 0x00

	// sheet type
	var st uint8 = 0x00

	biff8SheetName := UTF8toBIFF8UnicodeShort(sheetName)
	length := 6 + len(biff8SheetName)

	putVar(buffer, record, uint16(length))
	putVar(buffer, offset, ss, st)
	putVar(buffer, []byte(biff8SheetName))
}
