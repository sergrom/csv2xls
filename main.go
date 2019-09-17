package main

import (
	"bufio"
	"bytes"
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"math"
	"os"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

const OlePpsTypeRoot         = 5
const OlePpsTypeDir          = 1
const OlePpsTypeFile         = 2
const OleDataSizeSmall       = 0x1000
const OleLongIntSize         = 4
const OlePpsSize             = 0x80

type StringCollection struct {
	stringGrid         [][]string
	stringMap          map[string]int
	stringList         []string
	stringTotal        int
	stringUnique       int
}

func (sc *StringCollection) addRow(row []string)  {
	sc.stringGrid = append(sc.stringGrid, row)
	for _, str := range row {
		strToSave := UTF8toBIFF8UnicodeLong(str)
		if _, ok := sc.stringMap[strToSave]; !ok {
			sc.stringMap[strToSave] = sc.stringUnique
			sc.stringList = append(sc.stringList, strToSave)
			sc.stringUnique++
		}

		sc.stringTotal++
	}
}

type DataSectionItem struct {
	summary      uint32
	offset       uint32
	sType        uint32
	dataInt      uint32
	dataString   string
	dataLength   uint32
}

func main() {
	var CsvFileName string
	var CsvDelimiter string
	var XlsFileName string
	var Title string
	var Subject string
	var Creator string
	var Keywords string
	var Description string
	var LastModifiedBy string

	flag.StringVar(&CsvFileName, "csv-file-name", "", "The name of csv file that you want to convert to xls")
	flag.StringVar(&CsvDelimiter, "csv-delimiter", ";", "Delimiter on csv file")
	flag.StringVar(&XlsFileName, "xls-file-name", "","The name of result xls file")
	flag.StringVar(&Title, "title", "","Title property")
	flag.StringVar(&Subject, "subject", "","Subject property")
	flag.StringVar(&Creator, "creator", "","Creator property")
	flag.StringVar(&Keywords, "keywords", "","Keywords property")
	flag.StringVar(&Description, "description", "","Description property")
	flag.StringVar(&LastModifiedBy, "last-modified-by", "","LastModifiedBy property")
	flag.Parse()

	if CsvFileName == "" || XlsFileName == "" {
		flag.Usage()
		return
	}

	CsvDelimiterRune := ';'
	if CsvDelimiter != "" {
		if len(CsvDelimiter) > 1 {
			fmt.Println("Csv delimiter must be one character string")
			return
		}

		r, _ := utf8.DecodeRuneInString(CsvDelimiter)
		CsvDelimiterRune = r
	}

	var CreatedAtInt int64 = time.Now().Unix()
	var ModifiedAtInt int64 = time.Now().Unix()

	columnWidths := make(map[int]int, 0)
	//columnWidths[1] = 40 // parameter todo

	stringCollection := getStringCollectionFromCsvFile(CsvFileName, CsvDelimiterRune)

	wsArr := make([]Worksheet, 0)
	n := 0
	for i := 0; i<len(stringCollection.stringGrid); i += 65535 {
		wsName := "Worksheet"
		if n > 0 {
			wsName += strconv.Itoa(n)
		}

		last := i + 65535
		if last > len(stringCollection.stringGrid) {
			last = len(stringCollection.stringGrid)
		}
		wsArr = append(wsArr, Worksheet{wsName, stringCollection.stringGrid[i:last], columnWidths})
		n++
	}

	worksheetDatas := make([]string, 0)
	worksheetNames := make([]string, 0)
	for _, ws := range wsArr {
		worksheetDatas = append(worksheetDatas, ws.getData(&stringCollection))
		worksheetNames = append(worksheetNames, ws.Name)
	}

	worksheetSizes := make([]int, 0)
	for _, wsd := range worksheetDatas {
		worksheetSizes = append(worksheetSizes, len(wsd))
	}

	workbook := Workbook{worksheetSizes, worksheetNames, &stringCollection}

	var data strings.Builder
	data.WriteString(workbook.getWorksheetSizesData())

	for _, wsd := range worksheetDatas {
		data.WriteString(wsd)
	}

	rootPps := Pps{0, ascToUcs("Root Entry"), OlePpsTypeRoot, 0xFFFFFFFF, 0xFFFFFFFF, 1, "", 0, 0}
	workbookPps := Pps{1, ascToUcs("Workbook"), OlePpsTypeFile, 2, 3, 0xFFFFFFFF, data.String(), 0, 0}

	// TODO
	//documentSummaryInformationPps := Pps{2, fmt.Sprintf("%c%s", rune(5), ascToUcs("DocumentSummaryInformation")), OlePpsTypeFile, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, getDocumentSummaryInformation(), 0, 0}

	summaryInformation := getSummaryInformation(Title, Subject, Creator, Keywords, Description, LastModifiedBy, CreatedAtInt, ModifiedAtInt)
	summaryInformationPps := Pps{2, ascToUcs(fmt.Sprintf("%c%s", rune(5), "SummaryInformation")), OlePpsTypeFile, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF, summaryInformation, 0, 0}

	aList := []Pps{rootPps, workbookPps/*, TODO documentSummaryInformationPps*/, summaryInformationPps}

	iSBDcnt, iBBcnt, iPPScnt := calcSize(aList); // change types to uint32 TODO

	// Content of this buffer is result xls file
	resultBuffer := new(bytes.Buffer)

	saveHeader(resultBuffer, iSBDcnt, iBBcnt, iPPScnt)

	smallData := makeSmallData(resultBuffer, aList)
	aList[0].Data = smallData

	// Write BB
	saveBigData(resultBuffer, iSBDcnt, aList)

	// Write PPS
	savePps(resultBuffer, aList)

	// Write Big Block Depot and BDList and Adding Header informations
	saveBbd(resultBuffer, iSBDcnt, iBBcnt, iPPScnt)

	f, err := os.Create(XlsFileName)
	if err != nil {
		log.Fatal(err)
	}
	defer func() {
		err := f.Close()
		if err != nil {
			log.Fatal(err)
		}
	}()

	// Issue a `Sync` to flush writes to stable storage.
	err3 := f.Sync()
	if err3 != nil {
		log.Fatal(err3)
	}

	w := bufio.NewWriter(f)

	_, err2 := w.Write([]byte(resultBuffer.String()))
	if err2 != nil {
		log.Fatal(err)
	}

	// Use `Flush` to ensure all buffered operations have
	err4 := w.Flush()
	if err4 != nil {
		log.Fatal(err3)
	}
}

func saveBbd(buffer *bytes.Buffer, iSbdSize, iBsize, iPpsCnt uint32) {
	// Calculate Basic Setting
	var iBbCnt uint32 = 512 / OleLongIntSize
	var i1stBdL uint32 = (512 - 0x4C) / OleLongIntSize

	var iBdExL uint32 = 0
	iAll := iBsize + iPpsCnt + iSbdSize
	iAllW := iAll
	iBdCntW := uint32(math.Floor(float64(iAllW) / float64(iBbCnt)))
	if iAllW % iBbCnt > 0 {
		iBdCntW++
	}
	iBdCnt := uint32(math.Floor(float64(iAll + iBdCntW) / float64(iBbCnt)))
	if (iAllW + iBdCntW) % iBbCnt > 0 {
		iBdCnt++
	}
	// Calculate BD count
	if iBdCnt > i1stBdL {
		for {
			iBdExL++
			iAllW++
			iBdCntW = uint32(math.Floor(float64(iAllW) / float64(iBbCnt)))
			if iAllW % iBbCnt > 0 {
				iBdCntW++
			}
			iBdCnt = uint32(math.Floor(float64(iAllW + iBdCntW) / float64(iBbCnt)))
			if (iAllW + iBdCntW) % iBbCnt > 0 {
				iBdCnt++
			}
			if iBdCnt <= (iBdExL * iBbCnt + i1stBdL) {
				break
			}
		}
	}

	// Making BD
	// Set for SBD
	if iSbdSize > 0 {
		var i uint32
		for i = 0; i < (iSbdSize - 1); i++ {
			putVar(buffer, i + 1)
		}
		putVar(buffer, []byte("\xFE\xFF\xFF\xFF")) // uint32(-2)
	}

	// Set for B
	var i uint32
	for i = 0; i < (iBsize - 1); i++ {
		putVar(buffer, i + iSbdSize + 1)
	}
	putVar(buffer, []byte("\xFE\xFF\xFF\xFF"))

	// Set for PPS
	for i = 0; i < (iPpsCnt - 1); i++ {
		putVar(buffer, i + iSbdSize + iBsize + 1)
	}
	putVar(buffer, []byte("\xFE\xFF\xFF\xFF"))

	// Set for BBD itself ( 0xFFFFFFFD : BBD)
	for i = 0; i < iBdCnt; i++ {
		putVar(buffer, uint32(0xFFFFFFFD))
	}

	// Set for ExtraBDList
	for i = 0; i < iBdExL; i++ {
		putVar(buffer, uint32(0xFFFFFFFC))
	}

	// Adjust for Block
	if (iAllW + iBdCnt) % iBbCnt > 0 {
		iBlock := iBbCnt - ((iAllW + iBdCnt) % iBbCnt)
		for i = 0; i < iBlock; i++ {
			putVar(buffer, []byte("\xFF\xFF\xFF\xFF"))
		}
	}

	// Extra BDList
	if iBdCnt > i1stBdL {
		var iN, iNb uint32
		for i = i1stBdL; i < iBdCnt; i++{
			if iN >= (iBbCnt - 1) {
				iN = 0
				iNb++
				putVar(buffer, iAll + iBdCnt + iNb)
			}
			putVar(buffer, iBsize + iSbdSize + iPpsCnt + i)
			iN++
		}
		if (iBdCnt - i1stBdL) % (iBbCnt - 1) > 0 {
			iB := (iBbCnt - 1) - ((iBdCnt - i1stBdL) % (iBbCnt - 1))
			for i = 0; i < iB; i++ {
				putVar(buffer, []byte("\xFF\xFF\xFF\xFF"))
			}
		}
		putVar(buffer, []byte("\xFE\xFF\xFF\xFF"))
	}
}

func savePps(buffer *bytes.Buffer, raList []Pps) {
	// Save each PPS WK
	for _, pps := range raList {
		putVar(buffer, []byte(pps.getPpsWk())) // maybe it'll be better to change return type to []byte
	}
	// Adjust for Block
	iCnt := len(raList)
	iBCnt := 512 / OlePpsSize
	if iCnt % iBCnt > 0 {
		putVar(buffer, []byte(strings.Repeat("\x00", (iBCnt - (iCnt % iBCnt)) * OlePpsSize)))
	}
}

func saveBigData(buffer *bytes.Buffer, iStBlk uint32, raList []Pps) {
	// cycle through PPS's
	for i, _ := range raList {
		if raList[i].PpsType != OlePpsTypeDir {
			raList[i].Size = uint32(len(raList[i].Data))
			if raList[i].Size >= OleDataSizeSmall || (raList[i].PpsType == OlePpsTypeRoot && len(raList[i].Data) != 0) {
				putVar(buffer, []byte(raList[i].Data))

				if raList[i].Size % 512 > 0 {
					putVar(buffer, []byte(strings.Repeat("\x00", 512 - int(raList[i].Size) % 512)))
				}
				// Set For PPS
				raList[i].StartBlock = iStBlk
				iStBlk += uint32(math.Floor(float64(raList[i].Size) / 512))
				if raList[i].Size % 512 > 0 {
					iStBlk++
				}
			}
		}
	}
}

func makeSmallData(buffer *bytes.Buffer, raList []Pps) string {
	var smallData strings.Builder
	var iSmBlk uint32 = 0

	for i, _ := range raList {
		// Make SBD, small data string
		if raList[i].PpsType == OlePpsTypeFile {
			if raList[i].Size <= 0 {
				continue
			}

			if raList[i].Size < OleDataSizeSmall {
				iSmbCnt := uint32(math.Floor(float64(raList[i].Size) / 64))
				if raList[i].Size % 64 > 0 {
					iSmbCnt++
				}
				jB := iSmbCnt - 1
				var j uint32
				for j = 0; j < jB; j++ {
					putVar(buffer, j + iSmBlk + 1)
				}
				putVar(buffer, []byte("\xFE\xFF\xFF\xFF")) // uint32(-2)

				smallData.WriteString(raList[i].Data)
				if raList[i].Size % 64 > 0 {
					smallData.WriteString(strings.Repeat("\x00", 64 - int(raList[i].Size % 64)))
				}
				// Set for PPS
				raList[i].StartBlock = iSmBlk
				iSmBlk += iSmbCnt
			}
		}
	}

	iSbCnt := uint32(math.Floor(512.0 / OleLongIntSize))
	if iSmBlk % iSbCnt > 0 {
		iB := iSbCnt - (iSmBlk % iSbCnt)
		var i uint32
		for i = 0; i < iB; i++ {
			putVar(buffer, []byte("\xFF\xFF\xFF\xFF"))
		}
	}

	return smallData.String()
}

func saveHeader(buffer *bytes.Buffer, iSBDcnt, iBBcnt, iPPScnt uint32) {
	// Calculate Basic Setting
	var iBlCnt uint32 = 512 / OleLongIntSize
	var i1stBdL uint32 = (512 - 0x4C) / OleLongIntSize

	var iBdExL uint32 = 0
	iAll := uint32(iBBcnt + iPPScnt + iSBDcnt)
	iAllW := iAll
	iBdCntW := uint32(math.Floor(float64(iAllW) / float64(iBlCnt)))
	if iAllW % iBlCnt > 0 {
		iBdCntW++
	}
	iBdCnt := uint32(math.Floor(float64(iAll + iBdCntW) / float64(iBlCnt)))
	if (iAllW + iBdCntW) % iBlCnt > 0 {
		iBdCnt++
	}

	// Calculate BD count
	if iBdCnt > i1stBdL {
		for {
			iBdExL++
			iAllW++
			iBdCntW = uint32(math.Floor(float64(iAllW) / float64(iBlCnt)))
			if iAllW % iBlCnt > 0 {
				iBdCntW++
			}
			iBdCnt = uint32(math.Floor(float64(iAllW + iBdCntW) / float64(iBlCnt)))
			if (iAllW + iBdCntW) % iBlCnt > 0 {
				iBdCnt++
			}
			if iBdCnt <= (iBdExL * iBlCnt + i1stBdL) {
				break
			}
		}
	}

	// Save Header
	putVar(buffer,
		[]byte("\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\x00\x00\x00\x00"),
		uint16(0x3b),
		uint16(0x03),
		[]byte("\xFE\xFF"), // uint16(-2),
		uint16(9),
		uint16(6),
		uint16(0),
		[]byte("\x00\x00\x00\x00"),
		[]byte("\x00\x00\x00\x00"),
		iBdCnt,
	    iBBcnt + iSBDcnt,
	    uint32(0),
	    uint32(0x1000),
	)
	if iSBDcnt > 0 {
		putVar(buffer, uint32(0))
	} else {
		putVar(buffer, []byte("\xFE\xFF\xFF\xFF"))
	}
	putVar(buffer, iSBDcnt)

	// Extra BDList Start, Count
	if iBdCnt < i1stBdL {
		putVar(buffer,
			[]byte("\xFE\xFF\xFF\xFF"), // Extra BDList Start
			uint32(0), // Extra BDList Count
		)
	} else {
		putVar(buffer, iAll + iBdCnt, iBdExL)
	}

	// BDList
	var i uint32
	for i = 0; i < i1stBdL && i < iBdCnt; i++ {
		putVar(buffer, iAll + i)
	}
	if i < i1stBdL {
		jB := i1stBdL - i
		var j uint32
		for j = 0; j < jB; j++ {
			putVar(buffer, []byte("\xFF\xFF\xFF\xFF"))
		}
	}
}

func calcSize(aList []Pps) (uint32, uint32, uint32) {
	var iSBDcnt, iBBcnt, iPPScnt uint32 = 0, 0, 0

	iSBcnt := 0
	iCount := len(aList)
	for i := 0; i < iCount; i++ {
		if aList[i].PpsType == OlePpsTypeFile {
			aList[i].Size = uint32(len(aList[i].Data))

			if aList[i].Size < OleDataSizeSmall {
				iSBcnt += int(math.Floor(float64(aList[i].Size) / 64))
				if aList[i].Size % 64 > 0 {
					iSBcnt++
				}
			} else {
				iBBcnt += uint32(math.Floor(float64(aList[i].Size) / 512))
				if aList[i].Size % 512 > 0 {
					iBBcnt++
				}
			}
		}
	}

	iSlCnt := int(math.Floor(512 / OleLongIntSize))
	if (math.Floor(float64(iSBcnt) / float64(iSlCnt)) + float64(iSBcnt % iSlCnt)) > 0 {
		iSBDcnt = 1
	}

	iSmallLen := float64(iSBcnt) * 64
	iBBcnt += uint32(math.Floor(iSmallLen / 512))
	if int(iSmallLen) % 512 > 0 {
		iBBcnt++
	}
	iCnt := len(aList)
	iBdCnt := float64(512) / OlePpsSize
	iPPScnt = uint32(math.Floor(float64(iCnt) / iBdCnt))
	if iCnt % int(iBdCnt) > 0 {
		iPPScnt++
	}

	return iSBDcnt, iBBcnt, iPPScnt
}

func getSummaryInformation(title, subject, creator, keywords, description, lastModifiedBy string, created, modified int64) string {
	buffer := new(bytes.Buffer)

	// offset: 0; size: 2; must be 0xFE 0xFF (UTF-16 LE byte order mark)
	putVar(buffer, uint16(0xFFFE))
	// offset: 2; size: 2;
	putVar(buffer, uint16(0x0000))
	// offset: 4; size: 2; OS version
	putVar(buffer, uint16(0x0106))
	// offset: 6; size: 2; OS indicator
	putVar(buffer, uint16(0x0002))
	// offset: 8; size: 16
	putVar(buffer, uint32(0x00), uint32(0x00), uint32(0x00), uint32(0x00))
	// offset: 24; size: 4; section count
	putVar(buffer, uint32(0x0001))

	// offset: 28; size: 16; first section's class id: 02 d5 cd d5 9c 2e 1b 10 93 97 08 00 2b 2c f9 ae
	putVar(buffer, uint16(0x85E0), uint16(0xF29F), uint16(0x4FF9), uint16(0x1068), uint16(0x91AB), uint16(0x0008), uint16(0x272B), uint16(0xD9B3))
	// offset: 44; size: 4; offset of the start
	putVar(buffer, uint32(0x30))

	var dataSectionNumProps uint32 = 0
	dataSections := make([]DataSectionItem, 0)

	// CodePage : CP-1252
	dataSections = append(dataSections, DataSectionItem{0x01, 0, 0x02, 1252, "", 0})
	dataSectionNumProps++

	// Title
	if title != "" {
		dataSections = append(dataSections, DataSectionItem{0x02, 0, 0x1E, 0, title, uint32(len(title))})
		dataSectionNumProps++
	}

	// Subject
	if subject != "" {
		dataSections = append(dataSections, DataSectionItem{0x03, 0, 0x1E, 0, subject, uint32(len(subject))})
		dataSectionNumProps++
	}

	// Author (Creator)
	if creator != "" {
		dataSections = append(dataSections, DataSectionItem{0x04, 0, 0x1E, 0, creator, uint32(len(creator))})
		dataSectionNumProps++
	}

	// Keywords
	if keywords != "" {
		dataSections = append(dataSections, DataSectionItem{0x05, 0, 0x1E, 0, keywords, uint32(len(keywords))})
		dataSectionNumProps++
	}

	// Comments (Description)
	if description != "" {
		dataSections = append(dataSections, DataSectionItem{0x06, 0, 0x1E, 0, description, uint32(len(description))})
		dataSectionNumProps++
	}

	// Last Saved By (LastModifiedBy)
	if lastModifiedBy != "" {
		dataSections = append(dataSections, DataSectionItem{0x08, 0, 0x1E, 0, lastModifiedBy, uint32(len(lastModifiedBy))})
		dataSectionNumProps++
	}

	// Created Date/Time
	if created != 0 {
		dataSections = append(dataSections, DataSectionItem{0x0C, 0, 0x40, 0, localDateToOLE(created), 0})
		dataSectionNumProps++
	}

	// Modified Date/Time
	if modified != 0 {
		dataSections = append(dataSections, DataSectionItem{0x0D, 0, 0x40, 0, localDateToOLE(modified), 0})
		dataSectionNumProps++
	}

	// Security
	dataSections = append(dataSections, DataSectionItem{0x13, 0, 0x03, 0x00, "",  0})
	dataSectionNumProps++

	dataSectionSummary := new(bytes.Buffer)
	dataSectionContent := new(bytes.Buffer)
	dataSectionContentOffset := 8 + dataSectionNumProps * 8

	for _, dataSection := range dataSections {
		// Summary
		putVar(dataSectionSummary, dataSection.summary)
		// Offset
		putVar(dataSectionSummary, dataSectionContentOffset)
		// DataType
		putVar(dataSectionContent, dataSection.sType)
		// Data
		if dataSection.sType == 0x02 { // 2 byte signed integer
			putVar(dataSectionContent, dataSection.dataInt)
			dataSectionContentOffset += 8
		} else if dataSection.sType == 0x03 { // 4 byte signed integer
			putVar(dataSectionContent, dataSection.dataInt)
			dataSectionContentOffset += 8
		} else if dataSection.sType == 0x1E { // null-terminated string prepended by dword string length
			// Null-terminated string
			dataSection.dataString += "\x00"
			dataSection.dataLength++

			// Complete the string with null string for being a %4
			if (4 - dataSection.dataLength % 4) != 4 {
				dataSection.dataLength += 4 - dataSection.dataLength % 4
			}

			dataSection.dataString = dataSection.dataString + strings.Repeat("\x00", int(dataSection.dataLength) - len(dataSection.dataString))

			putVar(dataSectionContent, dataSection.dataLength)
			putVar(dataSectionContent, []byte(dataSection.dataString))

			dataSectionContentOffset += 8 + uint32(len(dataSection.dataString))
		} else if dataSection.sType == 0x40 { // Filetime (64-bit value representing the number of 100-nanosecond intervals since January 1, 1601)
			putVar(dataSectionContent, []byte(dataSection.dataString))
			dataSectionContentOffset += 4 + 8
		}
		// Data Type Not Used at the moment
	}
	// Now dataSectionContentOffset contains the size of the content

	// section header
	// offset: $secOffset; size: 4; section length
	//         + x  Size of the content (summary + content)
	putVar(buffer, dataSectionContentOffset)

	// offset: $secOffset+4; size: 4; property count
	putVar(buffer, dataSectionNumProps)

	// Section Summary
	putVar(buffer, dataSectionSummary.Bytes())

	// Section Content
	putVar(buffer, dataSectionContent.Bytes())

	return buffer.String()
}

func getDocumentSummaryInformation() string {
	return "" // TODO
}

func getGridAndStatFromCsvFile(csvFileName string) ([][]string, int, int, map[string]int) {
	grid := make([][]string, 0)
	stringTotal := 0
	stringUnique := 0
	stringTable := make(map[string]int, 0)

	f, err := os.Open(csvFileName)
	if err != nil {
		panic(fmt.Sprintf("Cannot read csv file \"%s\"", csvFileName))
	}
	defer f.Close()

	r := csv.NewReader(f)
	r.Comma = ';'

	for {
		record, err := r.Read()
		// Stop at EOF.
		if err == io.EOF {
			break
		}

		if err != nil {
			panic(err)
		}

		grid = append(grid, record)

		for _, value := range record {
			str := UTF8toBIFF8UnicodeLong(value)
			if _, ok := stringTable[str]; !ok {
				stringTable[str] = stringUnique
				stringUnique++
			}

			stringTotal++
		}
	}

	return grid, stringTotal, stringUnique, stringTable
}

func getStringCollectionFromCsvFile(csvFileName string, delimiter rune) StringCollection {
	sc := StringCollection{make([][]string, 0), make(map[string]int,0), make([]string, 0), 0,0}

	f, err := os.Open(csvFileName)
	if err != nil {
		panic(fmt.Sprintf("Cannot read csv file \"%s\"", csvFileName))
	}
	defer f.Close()

	r := csv.NewReader(f)
	r.Comma = delimiter
	r.LazyQuotes = true

	for {
		record, err := r.Read()
		// Stop at EOF.
		if err == io.EOF {
			break
		}

		if err != nil {
			panic(err)
		}

		sc.addRow(record)
	}

	return sc
}