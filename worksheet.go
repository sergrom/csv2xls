package main

import (
	"bytes"
)

type Worksheet struct {
	Name string
	Grid [][]string
	ColumnWidths map[int]int
}

func (ws *Worksheet) getName() string {
	return ws.Name
}

func (ws *Worksheet) getData(stringCollection *StringCollection) string {
	buf := new(bytes.Buffer)

	maxColIdx := 0
	for _, row := range(ws.Grid) {
		maxColIdx = Max(maxColIdx, len(row) -1)
	}

	// Write BOF record
	ws.storeBof(buf)

	// Write PRINTHEADERS
	ws.writePrintHeaders(buf)

	// Write PRINTGRIDLINES
	ws.writePrintGridlines(buf)

	// Write GRIDSET
	ws.writeGridset(buf)

	columnInfo := make([][]uint16, 0);
	for i:=0; i <= maxColIdx; i++ {
		w := 10
		if value, ok := ws.ColumnWidths[i]; ok {
			w = value
		}
		var info = []uint16{uint16(i), uint16(i), uint16(w), 15, 0, 0}
		columnInfo = append(columnInfo, info)
	}

	// Write GUTS
	ws.writeGuts(buf, columnInfo)
	// Write DEFAULTROWHEIGHT
	ws.writeDefaultRowHeight(buf)
	// Write WSBOOL
	ws.writeWsbool(buf)
	// Write horizontal and vertical page breaks
	ws.writeBreaks(buf)
	// Write page header
	ws.writeHeader(buf)
	// Write page footer
	ws.writeFooter(buf)
	// Write page horizontal centering
	ws.writeHcenter(buf)
	// Write page vertical centering
	ws.writeVcenter(buf)
	// Write left margin
	ws.writeMarginLeft(buf)
	// Write right margin
	ws.writeMarginRight(buf)
	// Write top margin
	ws.writeMarginTop(buf)
	// Write bottom margin
	ws.writeMarginBottom(buf)
	// Write page setup
	ws.writeSetup(buf)
	// Write sheet protection
	ws.writeProtect(buf)
	// Write SCENPROTECT
	ws.writeScenProtect(buf)
	// Write OBJECTPROTECT
	ws.writeObjectProtect(buf)
	// Write sheet password
	ws.writePassword(buf)
	// Write DEFCOLWIDTH record
	ws.writeDefcol(buf)

	// Write the COLINFO records if they exist
	if len(columnInfo) != 0 {
		colcount := len(columnInfo)
		for i :=0; i<colcount; i++ {
			ws.writeColinfo(buf, columnInfo[i])
		}
	}

	// Write sheet dimensions
	var firstRowIndex uint32 = 0
	lastRowIndex := uint32(len(ws.Grid))
	var firstColumnIndex uint16 = 1
	var lastColumnIndex uint16 = uint16(maxColIdx) + 1

	ws.writeDimensions(buf, firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex)

	// Write Cells
	for rowIdx, rows := range(ws.Grid) {
		for columnIdx, cValue := range(rows) {
			if rowIdx > 65535 || columnIdx > 255 {
				panic("Rows or columns overflow! Excel5 has limit to 65535 rows and 255 columns. Use XLSX instead.")
			}

			// Write cell value
			if cValue == "" {
				ws.writeBlank(buf, rowIdx, columnIdx, 15)
			} else {
				ws.writeString(buf, rowIdx, columnIdx, cValue, 15, stringCollection)
			}
		}
	}

	// Append
	ws.writeMsoDrawing(buf)

	// Write WINDOW2 record
	ws.writeWindow2(buf)

	// Write PLV record
	ws.writePageLayoutView(buf)

	// Write ZOOM record
	ws.writeZoom(buf)

	// Write SELECTION record
	ws.writeSelection(buf)

	// Write MergedCellsTable Record
	ws.writeMergedCells(buf)

	ws.writeDataValidity(buf)
	ws.writeSheetLayout(buf)

	// Write SHEETPROTECTION record
	ws.writeSheetProtection(buf)
	ws.writeRangeProtection(buf)

	ws.storeEof(buf)

	return buf.String()
}

func (ws *Worksheet) storeBof(buffer *bytes.Buffer) {
	var bType uint16 = 0x0010

	var record uint16 = 0x0809 // Record identifier    (BIFF5-BIFF8)
	var length uint16 = 0x0010

	var build uint16 = 0x0DBB //    Excel 97
	var year uint16 = 0x07CC //    Excel 97
	var version uint16 = 0x0600 //    BIFF8

	putVar(buffer, record, length, version, bType, build, year)

	// by inspection of real files, MS Office Excel 2007 writes the following
	putVar(buffer, uint32(0x000100D1), uint32(0x00000406))
}

func (ws *Worksheet) writePrintHeaders(buffer *bytes.Buffer) {
	var record uint16 = 0x002a // Record identifier
	var length uint16 = 0x0002 // Bytes to follow

	var fPrintRwCol uint16 = 0 // Boolean flag

	putVar(buffer, record, length, fPrintRwCol)
}

func (ws *Worksheet) writePrintGridlines(buffer *bytes.Buffer) {
	var record uint16 = 0x002b // Record identifier
	var length uint16 = 0x0002 // Bytes to follow

	var fPrintGrid uint16 = 0 // Boolean flag

	putVar(buffer, record, length, fPrintGrid)
}

func (ws *Worksheet) writeGridset(buffer *bytes.Buffer) {
	var record uint16 = 0x0082 // Record identifier
	var length uint16 = 0x0002 // Bytes to follow

	var fGridSet uint16 = 1 // Boolean flag

	putVar(buffer, record, length, fGridSet)
}

func (ws *Worksheet) writeGuts(buffer *bytes.Buffer, columnInfo [][]uint16) {
	var record uint16 = 0x0080 // Record identifier
	var length uint16 = 0x0008 // Bytes to follow

	var dxRwGut uint16 = 0x0000 // Size of row gutter
	var dxColGut uint16 = 0x0000 // Size of col gutter

	// determine maximum row outline level
	var maxRowOutlineLevel uint16 = 0

	var col_level uint16 = 0

	// Calculate the maximum column outline level. The equivalent calculation
	// for the row outline level is carried out in writeRow().
	colcount := len(columnInfo)
	for i := 0; i < colcount; i++ {
		col_level = MaxUInt16(columnInfo[i][5], col_level)
	}

	// Set the limits for the outline levels (0 <= x <= 7).
	col_level = MaxUInt16(0, MinUInt16(col_level, 7))

	if col_level != 0 {
		col_level++
	}

	putVar(buffer, record, length)
	putVar(buffer, dxRwGut, dxColGut, maxRowOutlineLevel, col_level)
}

func (ws *Worksheet) writeDefaultRowHeight(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeWsbool(buffer *bytes.Buffer) {
	var record uint16 = 0x0081 // Record identifier
	var length uint16 = 0x0002 // Bytes to follow
	var grbit uint16 = 0x0000

	// Set the option flags
	grbit |= 0x0001 // Auto page breaks visible
	grbit |= 0x0040 // Outline summary below
	grbit |= 0x0080 // Outline summary right
	grbit |= 0x0400 // Outline symbols displayed

	putVar(buffer, record, length, grbit)
}

func (ws *Worksheet) writeBreaks(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeHeader(buffer *bytes.Buffer) {
	var record uint16 = 0x0014 // Record identifier
	recordData := UTF8toBIFF8UnicodeLong("")
	length := uint16(len(recordData))

	putVar(buffer, record, length, []byte(recordData))
}

func (ws *Worksheet) writeFooter(buffer *bytes.Buffer) {
	var record uint16 = 0x0015 // Record identifier
	recordData := UTF8toBIFF8UnicodeLong("")
	length := uint16(len(recordData))

	putVar(buffer, record, length, []byte(recordData))
}

func (ws *Worksheet) writeHcenter(buffer *bytes.Buffer) {
	var record uint16 = 0x0083 // Record identifier
	var length uint16 = 0x0002 // Bytes to follow

	var fHCenter uint16 = 0 // Horizontal centering

	putVar(buffer, record, length, fHCenter)
}

func (ws *Worksheet) writeVcenter(buffer *bytes.Buffer) {
	var record uint16 = 0x0084 // Record identifier
	var length uint16 = 0x0002 // Bytes to follow

	var fVCenter uint16 = 0 // Horizontal centering

	putVar(buffer, record, length, fVCenter)
}

func (ws *Worksheet) writeMarginLeft(buffer *bytes.Buffer) {
	var record uint16 = 0x0026 // Record identifier
	var length uint16 = 0x0008 // Bytes to follow

	margin := 0.7 // Margin in inches

	putVar(buffer, record, length, margin)
}

func (ws *Worksheet) writeMarginRight(buffer *bytes.Buffer) {
	var record uint16 = 0x0027 // Record identifier
	var length uint16 = 0x0008 // Bytes to follow

	margin := 0.7 // Margin in inches

	putVar(buffer, record, length, margin)
}

func (ws *Worksheet) writeMarginTop(buffer *bytes.Buffer) {
	var record uint16 = 0x0028 // Record identifier
	var length uint16 = 0x0008 // Bytes to follow

	margin := 0.75 // Margin in inches

	putVar(buffer, record, length, margin)
}

func (ws *Worksheet) writeMarginBottom(buffer *bytes.Buffer) {
	var record uint16 = 0x0029 // Record identifier
	var length uint16 = 0x0008 // Bytes to follow

	margin := 0.75 // Margin in inches

	putVar(buffer, record, length, margin)
}

func (ws *Worksheet) writeSetup(buffer *bytes.Buffer) {
	var record uint16 = 0x00A1 // Record identifier
	var length uint16 = 0x0022 // Number of bytes to follow

	var iPaperSize uint16 = 1 // Paper size

	var iScale uint16 = 100 // Print scaling factor

	var iPageStart uint16 = 0x01 // Starting page number
	var iFitWidth uint16 = 1 // Fit to number of pages wide
	var iFitHeight uint16 = 1 // Fit to number of pages high
	var iRes uint16 = 0x0258 // Print resolution
	var iVRes uint16 = 0x0258 // Vertical print resolution

	numHdr := 0.3 // Header Margin

	numFtr := 0.3 // Footer Margin
	iCopies := uint16(0x01) // Number of copies

	var fLeftToRight uint16 = 0x0 // Print over then down

	// Page orientation
	var fLandscape uint16 = 0x1

	var fNoPls uint16 = 0x0 // Setup not read from printer
	var fNoColor uint16 = 0x0 // Print black and white
	var fDraft uint16 = 0x0 // Print draft quality
	var fNotes uint16 = 0x0 // Print notes
	var fNoOrient uint16 = 0x0 // Orientation not set
	var fUsePage uint16 = 0x0 // Use custom starting page

	grbit := fLeftToRight
	grbit |= fLandscape << 1
	grbit |= fNoPls << 2
	grbit |= fNoColor << 3
	grbit |= fDraft << 4
	grbit |= fNotes << 5
	grbit |= fNoOrient << 6
	grbit |= fUsePage << 7

	putVar(buffer, record, length)
	putVar(buffer, iPaperSize, iScale, iPageStart, iFitWidth, iFitHeight, grbit, iRes, iVRes)
	putVar(buffer, numHdr, numFtr)
	putVar(buffer, iCopies)
}

func (ws *Worksheet) writeProtect(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeScenProtect(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeObjectProtect(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writePassword(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeDefcol(buffer *bytes.Buffer) {
	var defaultColWidth uint16 = 8

	var record uint16 = 0x0055 // Record identifier
	var length uint16 = 0x0002 // Number of bytes to follow

	putVar(buffer, record, length, defaultColWidth)
}

func (ws *Worksheet) writeColinfo(buffer *bytes.Buffer, columnInfo []uint16) {
	var colFirst, colLast, grbit, level uint16
	var coldx uint16 = 10
	var xfIndex uint16 = 15

	for i, v := range(columnInfo) {
		switch i {
		case 0:
			colFirst = v
			break
		case 1:
			colLast = v
			break
		case 2:
			coldx = v
			break
		case 3:
			xfIndex = v
			break
		case 4:
			grbit = v
			break
		case 5:
			level = v
			break
		}
	}
	var record uint16 = 0x007D // Record identifier
	var length uint16 = 0x000C // Number of bytes to follow

	coldx *= 256 // Convert to units of 1/256 of a char

	ixfe := xfIndex
	var reserved uint16 = 0x0000 // Reserved

	level = MaxUInt16(0, MinUInt16(level, 7))
	grbit |= level << 8;

	putVar(buffer, record, length)
	putVar(buffer, colFirst, colLast, coldx, ixfe, grbit, reserved)
}

func (ws *Worksheet) writeDimensions(buffer *bytes.Buffer, firstRowIndex uint32, lastRowIndex uint32, firstColumnIndex uint16, lastColumnIndex uint16) {
	var record uint16 = 0x0200 // Record identifier
	var length uint16 = 0x000E

	putVar(buffer, record, length, firstRowIndex, lastRowIndex + 1, firstColumnIndex, lastColumnIndex + 1, uint16(0x0000))
}

func (ws *Worksheet) writeBlank(buffer *bytes.Buffer, rowIdx int, columnIdx int, xfIndex int) {
	var record uint16 = 0x0201 // Record identifier
	var length uint16 = 0x0006 // Number of bytes to follow

	putVar(buffer, record, length)
	putVar(buffer, uint16(rowIdx), uint16(columnIdx), uint16(xfIndex))
}

func (ws *Worksheet) writeString(buffer *bytes.Buffer, rowIdx int, columnIdx int, cValue string, xfIndex int, stringCollection *StringCollection) {
	var record uint16 = 0x00FD // Record identifier
	var length uint16 = 0x000A // Bytes to follow

	cValue = UTF8toBIFF8UnicodeLong(cValue)

	if strTabVal, ok := stringCollection.stringMap[cValue]; ok {
		putVar(buffer, record, length)
		putVar(buffer, uint16(rowIdx), uint16(columnIdx), uint16(xfIndex), uint32(strTabVal))
		return
	}

	panic("Something happened wrong")
}

func (ws *Worksheet) writeMsoDrawing(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeWindow2(buffer *bytes.Buffer) {
	var record uint16 = 0x023E // Record identifier
	var length uint16 = 0x0012

	var rwTop uint16 = 0x0000 // Top row visible in window
	var colLeft uint16 = 0x0000 // Leftmost column visible in window

	// The options flags that comprise $grbit
	var fDspFmla uint16 = 0 // 0 - bit
	var fDspGrid uint16 = 1 // 1
	var fDspRwCol uint16 = 1 // 2
	var fFrozen uint16 = 0 // 3
	var fDspZeros uint16 = 1 // 4
	var fDefaultHdr uint16 = 1 // 5
	var fArabic uint16 = 0 // 6
	var fDspGuts uint16 = 1 // 7
	var fFrozenNoSplit uint16 = 0 // 0 - bit
	// no support in PhpSpreadsheet for selected sheet, therefore sheet is only selected if it is the active sheet
	var fSelected uint16 = 1
	var fPaged uint16 = 1 // 2
	var fPageBreakPreview uint16 = 0

	grbit := fDspFmla;
	grbit |= fDspGrid << 1
	grbit |= fDspRwCol << 2
	grbit |= fFrozen << 3
	grbit |= fDspZeros << 4
	grbit |= fDefaultHdr << 5
	grbit |= fArabic << 6
	grbit |= fDspGuts << 7
	grbit |= fFrozenNoSplit << 8
	grbit |= fSelected << 9
	grbit |= fPaged << 10
	grbit |= fPageBreakPreview << 11

	putVar(buffer, record, length)
	putVar(buffer, grbit, rwTop, colLeft)

	var rgbHdr uint16 = 0x0040; // Row/column heading and gridline color index
	var zoom_factor_page_break uint16 = 0;
	var zoom_factor_normal uint16 = 100;

	putVar(buffer, rgbHdr, uint16(0x0000), zoom_factor_page_break, zoom_factor_normal, uint32(0x00000000))
}

func (ws *Worksheet) writePageLayoutView(buffer *bytes.Buffer) {
	var record uint16 = 0x088B // Record identifier
	var length uint16 = 0x0010 // Bytes to follow

	var rt uint16 = 0x088B // 2
	var grbitFrt uint16 = 0x0000 // 2
	var wScalvePLV uint16 = 100 //$this->phpSheet->getSheetView()->getZoomScale(); // 2

	var fPageLayoutView uint16 = 0
	var fRulerVisible uint16 = 0
	var fWhitespaceHidden uint16 = 0;

	grbit := fPageLayoutView // 2
	grbit |= fRulerVisible << 1
	grbit |= fWhitespaceHidden << 3

	putVar(buffer, record, length)
	putVar(buffer,  rt, grbitFrt, uint32(0x00000000), uint32(0x00000000), wScalvePLV, grbit)
}

func (ws *Worksheet) writeZoom(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeSelection(buffer *bytes.Buffer) {
	var record uint16 = 0x001D // Record identifier
	var length uint16 = 0x000F // Number of bytes to follow

	putVar(buffer, record, length)
	putVar(buffer, uint8(3), uint16(0), uint16(0), uint16(0), uint16(1), uint16(0), uint16(0), uint8(0), uint8(0))
}

func (ws *Worksheet) writeMergedCells(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeDataValidity(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeSheetLayout(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) writeSheetProtection(buffer *bytes.Buffer) {
	// record identifier
	var record uint16 = 0x0867
	var length uint16 = 23

	// prepare options
	var options uint16 = 32767

	putVar(buffer, record, length)
	putVar(buffer, uint16(0x0867), uint32(0x0000), uint32(0x0000), uint8(0x00), uint32(0x01000200), uint32(0xFFFFFFFF), options, uint16(0x0000))
}

func (ws *Worksheet) writeRangeProtection(buffer *bytes.Buffer) {
	// empty
}

func (ws *Worksheet) storeEof(buffer *bytes.Buffer) {
	var record uint16 = 0x000A; // Record identifier
	var length uint16 = 0x0000; // Number of bytes to follow

	putVar(buffer, record, length)
}

