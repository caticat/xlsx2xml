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
	// parse config
	flag.Parse()
	pathIn, pathOut, fmtEnable, fmtOut := loadConfig()
	fmt.Println(fmtEnable, fmtOut)

	// processing
	for _, pathSrc := range getFilelist(pathIn) {
		// Args to store tmp data
		vKey, vDesc, vData := []string{}, []string{}, []string{}

		// Analyze data from xlsx
		analyzeXlsx(pathSrc, &vKey, &vDesc, &vData)

		// Write data to xml
		if len(vKey) > 0 {
			pathTar := strings.TrimSuffix(pathOut+getRelativeDir(pathIn, pathSrc), ".xlsx") + ".xml"
			writeXml(pathTar, vKey, vDesc, vData)
		}
	}

	return
}

func loadConfig() (string, string, bool, string) {
	cfg := *conf.LoadConfig(config)
	format_file_enable, _ := strconv.ParseBool(cfg.Read("format_file", "enable"))
	return cfg.Read("path", "in"), cfg.Read("path", "out"), format_file_enable, cfg.Read("format_file", "out")
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

func analyzeXlsx(pathSrc string, keyV, descV, dataV *[]string) {
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
	}

	// organizing data
	keyAllV := []string{}
	for _, sheet := range xlFile.Sheets {
		for x, row := range sheet.Rows {
			switch x {
			case XTitle:
			case XType:
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
	}
}

func writeXml(pathTar string, keyV, descV, dataV []string) {
	// check output floder
	dirname := getDirname(pathTar)
	if !isPathExist(dirname) {
		createPath(dirname)
	}

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
	if _, err := writer.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"); err != nil {
		panic(err)
	}
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

func writeFormat() {

}
