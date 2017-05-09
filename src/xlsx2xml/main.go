package main

import (
	"bufio"
	"conf"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

var (
	config = flag.String("config", "config.ini", "General configuraion file.")
)

// every valuable columns
var (
	Export_Usage = map[string]bool{"Both": true, "Server": true}
)

// rows of xlsx
const (
	XTitle = iota // 标题
	XType         // 数据类型
	XDesc         // 字段描述
	XUsage        // 使用范围
	XName         // 字段名
	XData         // 具体数据
)

func main() {
	defer func(timeBegin int64) {
		log("done. cost %d seconds\n", time.Now().Unix()-timeBegin)
		time.Sleep(time.Second * 3)
	}(time.Now().Unix())

	// parse config
	log("loading config\n")
	flag.Parse()
	pathIn, pathOut, fmtEnable, fmtOut := loadConfig()
	log("path xlsx:%s\n", pathIn)
	log("path xml:%s\n", pathOut)
	if fmtEnable {
		log("path fmt:%s\n", fmtOut)
	}

	// processing
	for _, pathSrc := range getFilelist(pathIn) {
		log("parsing file %s\n", pathSrc)

		// Args to store tmp data
		vKey, vDesc, vData, vType := []string{}, []string{}, []string{}, []string{}

		// Analyze data from xlsx
		analyzeXlsx(pathSrc, &vKey, &vDesc, &vData, &vType)

		// Write data to xml
		if len(vKey) > 0 {
			pathTar := strings.TrimSuffix(pathOut+getRelativeDir(pathIn, pathSrc), ".xlsx") + ".xml"
			writeXml(pathTar, vKey, vDesc, vData)

			if fmtEnable {
				pathTar = strings.TrimSuffix(fmtOut+getRelativeDir(pathIn, pathSrc), ".xlsx") + ".txt"
				writeFormat(pathTar, vKey, vDesc, vType)
			}
		}
	}

	//	bufio.NewReader(os.Stdin).ReadLine()
	return
}

func log(format string, args ...interface{}) {
	if len(args) == 0 {
		fmt.Printf(format)
	} else {
		fmt.Printf(format, args)
	}
}

func loadConfig() (string, string, bool, string) {
	cfg := *conf.LoadConfig(config)
	format_file_enable, _ := strconv.ParseBool(cfg.Read("format_file", "enable"))
	return fixPath(cfg.Read("path", "in")), fixPath(cfg.Read("path", "out")), format_file_enable, fixPath(cfg.Read("format_file", "out"))
}

func fixPath(path string) string {
	return strings.Replace(path, "\\", "/", -1)
}

// get all files in the path
func getFilelist(path string) []string {
	fileV := []string{}
	filepath.Walk(path, func(path string, f os.FileInfo, err error) error {
		if f == nil {
			panic(err)
		}
		if f.IsDir() {
			return nil
		}
		fileV = append(fileV, path)
		return nil
	})
	return fileV
}

func getRelativeDir(base, full string) string {
	full = strings.Replace(full, "\\", "/", -1)
	return full[strings.Index(full, base)+len(base):]
}

func getDirname(path string) string {
	if pos := strings.LastIndex(path, "/"); pos == -1 {
		return path
	} else {
		return path[:pos]
	}
}

func isPathExist(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}

func createPath(path string) {
	if err := os.MkdirAll(path, os.ModePerm); err != nil {
		panic(err)
	}
}

func analyzeXlsx(pathSrc string, keyV, descV, dataV, typeV *[]string) {
	// open file
	xlFile, err := xlsx.OpenFile(pathSrc)
	if err != nil {
		panic(err)
	}

	// get valid columns
	mValid := make(map[int]bool)
	for _, sheet := range xlFile.Sheets {
		if len(sheet.Rows) < XData {
			return
		}
		for y, cell := range sheet.Rows[XUsage].Cells {
			text, _ := cell.String()
			if _, ok := Export_Usage[text]; ok {
				mValid[y] = true
			}
		}
		break
	}

	// organizing data
	keyAllV := []string{}
	for _, sheet := range xlFile.Sheets {
		for x, row := range sheet.Rows {
			switch x {
			case XTitle:
			case XType:
				for y, cell := range row.Cells {
					if _, ok := mValid[y]; !ok {
						continue
					}
					text, _ := cell.String()
					*typeV = append(*typeV, text)
				}
			case XDesc:
				for y, cell := range row.Cells {
					if _, ok := mValid[y]; !ok {
						continue
					}
					text, _ := cell.String()
					*descV = append(*descV, text)
				}
			case XUsage:
			case XName:
				for y, cell := range row.Cells {
					text, _ := cell.String()
					keyAllV = append(keyAllV, text)
					if _, ok := mValid[y]; !ok {
						continue
					}
					*keyV = append(*keyV, text)
				}
			default:
				dataS := "\t<data"
				for y, cell := range row.Cells {
					if _, ok := mValid[y]; !ok {
						continue
					} else if y >= len(keyAllV) {
						break
					}
					text, _ := cell.String()
					dataS = fmt.Sprintf("%s %s=\"%s\"", dataS, (keyAllV)[y], text) // If slice can trans to array
				}
				*dataV = append(*dataV, dataS+" />\n")
			}
		}
		break
	}
}

func writeXml(pathTar string, keyV, descV, dataV []string) {
	// check output floder
	dirname := getDirname(pathTar)
	if !isPathExist(dirname) {
		createPath(dirname)
	}

	// write file
	file, err := os.OpenFile(pathTar, os.O_CREATE|os.O_RDWR, 0664)
	if err != nil {
		panic(err)
	}
	defer file.Close()
	if err := file.Truncate(0); err != nil {
		panic(err)
	}
	writer := bufio.NewWriter(file)
	defer writer.Flush()

	// write head
	if _, err := writer.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"); err != nil {
		panic(err)
	}

	// write comment
	if _, err := writer.WriteString("<!-- "); err != nil {
		panic(err)
	}
	for k, v := range keyV {
		if k >= len(descV) {
			break
		}
		if _, err := writer.WriteString(fmt.Sprintf("%s=%s ", v, descV[k])); err != nil {
			panic(err)
		}
	}
	if _, err := writer.WriteString("-->\n"); err != nil {
		panic(err)
	}

	// write data
	if _, err := writer.WriteString("<root>\n"); err != nil {
		panic(err)
	}
	for _, v := range dataV {
		_, err := writer.WriteString(v)
		if err != nil {
			panic(err)
		}
	}
	if _, err := writer.WriteString("</root>\n"); err != nil {
		panic(err)
	}
}

func writeFormat(pathTar string, keyV, descV, typeV []string) {
	// check output floder
	dirname := getDirname(pathTar)
	if !isPathExist(dirname) {
		createPath(dirname)
	}

	// get arr length
	length := func(args ...int) int {
		if len(args) == 0 {
			panic("invalid args")
		}
		arg := args[0]
		for _, v := range args {
			if v < arg {
				arg = v
			}
		}
		return arg
	}(len(keyV), len(descV), len(typeV))

	// write data
	file, err := os.OpenFile(pathTar, os.O_CREATE|os.O_RDWR, 0664)
	if err != nil {
		panic(err)
	}
	defer file.Close()
	if err := file.Truncate(0); err != nil {
		panic(err)
	}
	writer := bufio.NewWriter(file)
	defer writer.Flush()

	for i := 0; i < length; i++ {
		if _, err := writer.WriteString(fmt.Sprintf("\t%s\t%s\t`xml:\"%s,attr\"`\t//%s\n", keyV[i], typeV[i], keyV[i], descV[i])); err != nil {
			panic(err)
		}
	}
}
