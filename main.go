package main

import (
	"bytes"
	"io/ioutil"
	"log"
	"fmt"
	"strings"
	"os"
	"path/filepath"
	"regexp"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/ledongthuc/pdf"
)

var cna, lmm string

func main () {

	// remove existing output files
	if _, err := os.Stat("./files/outputCNA.txt"); !os.IsNotExist(err) {
		err = os.Remove("./files/outputCNA.txt")
		if err != nil {
			log.Fatal(err)
		}
	}

	if _, err := os.Stat("./files/outputLMM.txt"); !os.IsNotExist(err) {
		err = os.Remove("./files/outputLMM.txt")
		if err != nil {
			log.Fatal(err)
		}
	}

	// read file names
	files, err := ioutil.ReadDir("./files/")
	if err != nil {
		log.Fatal(err)
	}

	// Read from xlsx or pdf
	for _, fi := range files {
		if filepath.Ext(fi.Name()) == ".xlsx" {
			readFileXLSX(fi.Name())
		}
		if filepath.Ext(fi.Name()) == ".pdf" {
			cna, err = readFilePDF(fi.Name())
			if err != nil {
				panic(err)
			}
			cna = cna[102:len(cna)-13]
			reg := regexp.MustCompile("(Mon|Tue|Wed|Thur|Thr|Fri|Sat|Sun)[0-9]{1,2} [a-zA-Z]{3} [0-9]{4,6}")
			cna = reg.ReplaceAllString(cna, "")
			cna = strings.Replace(cna, ";,", ",", -1)
			cna = strings.Replace(cna, ";", ",", -1)
			cna = strings.Replace(cna, ".", ",", -1)
			cna = strings.Replace(cna, ",,", ",", -1)
			cna = strings.Replace(cna, "  ", " ", -1)
			cna = strings.Replace(cna, ", ", ",", -1)
			cna = strings.Replace(cna, ",", ", ", -1)

		}
	}

	// Write to a file CNA
	if len(cna) > 0 {
		foc, err := os.Create("./files/outputCNA.txt")
		if err != nil {
			log.Fatal(err)
		}
		defer foc.Close()

		_, err = foc.WriteString(cna)
		if err != nil {
			log.Fatal(err)
		}
		foc.Sync()
	}

	if len(lmm) > 0 {
		// Write to a file LMM
		fol, err := os.Create("./files/outputLMM.txt")
		if err != nil {
			log.Fatal(err)
		}
		defer fol.Close()

		_, err = fol.WriteString(lmm)
		if err != nil {
			log.Fatal(err)
		}
		fol.Sync()
	}
}

func readFileXLSX(s string) {
	xlsx, err := excelize.OpenFile("./files/" + s)
	if err != nil {
		log.Fatal(err)
	}

	// Detect column with lots of string
	max, maxCol := 0, 0
	rows := xlsx.GetRows(xlsx.GetSheetName(1))
		for _, row := range rows {
			for i := 0; i < len(row); i++ {
				if len(row[i]) > max {
					max, maxCol = len(row[i]), i
				}
			}
		}

	mode := "CNA"

	// Get all the rows in the Sheet1
	rows = xlsx.GetRows(xlsx.GetSheetName(1))
	for i, row := range rows {
		if i == 0 { continue }
		// Index out of range protection
		if len(row) > maxCol+1 {
			if (strings.Trim(row[maxCol + 1], " ") == "CNA" || strings.Trim(row[maxCol + 1], " ") == "CAN") && len(row[maxCol]) > 5 {
				mode = "CNA"
			}
			if strings.Trim(row[maxCol + 1], " ") == "LMM" && len(row[maxCol]) > 5 {
				mode = "LMM"
			}
		}

		if row[maxCol] != "" && len(row[maxCol]) > 14 {
			s := row[maxCol]
			s = strings.Replace(s, "\n", ", ", -1)
			s = strings.Trim(s, " ")
			s = strings.Replace(s, " ,", ",", -1)
			s = strings.Replace(s, ",,", ",", -1)
			if s[len(s) - 1] == ',' { s = s[:len(s) - 1] }
			if s == "," || len(s) == 0  { continue }
			if mode == "CNA" {
				cna = cna + s + ", "
			}
			if mode == "LMM" {
				lmm = lmm + s + ", "
			}
		}
	}
	cna = strings.Replace(cna, ";,", ",", -1)
	cna = strings.Replace(cna, ";", ",", -1)
	cna = strings.Replace(cna, ".", ",", -1)
	cna = strings.Replace(cna, ",,", ",", -1)
	lmm = strings.Replace(lmm, ";,", ",", -1)
	lmm = strings.Replace(lmm, ";", ",", -1)
	lmm = strings.Replace(lmm, ".", ",", -1)
	lmm = strings.Replace(lmm, ",,", ",", -1)

	cna = cna + "\n"
	if len(lmm) > 0 {
		lmm = lmm + "\n"
	}
	fmt.Print()
}

func readFilePDF(s string) (string, error) {
	f, r, err := pdf.Open("./files/" + s)
	// remember close file
	defer f.Close()
	if err != nil {
		return "", err
	}
	var buf bytes.Buffer
	b, err := r.GetPlainText()
	if err != nil {
		return "", err
	}
	buf.ReadFrom(b)
	return buf.String(), nil
}
