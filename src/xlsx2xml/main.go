package main

import (
	"bufio"
	"conf"
	"errors"
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
	config = flag.String("config", "xlsx2xml.ini", "General configuraion file.")
)

// every valuable columns
var (
	Export_Usage = map[string]bool{"Both": true, "Server": true}
)

// type conversion
var (
	Type_Conversion = map[string]string{"int": "int64" /*, "string": "string"*/}
)

// value conversion
var (
	Value_Conversion = map[string]func(string) bool{
		"int": func(val string) bool {
			_, err := strconv.ParseInt(val, 10, 64)
			return err == nil
		}}
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
	var err error
	var validFileCount int
	defer func(timeBegin int64) {
		if err == nil {
			log("%d files converted.\ncost %d seconds.\ndone.\n", validFileCount, time.Now().Unix()-timeBegin)
			time.Sleep(time.Second * 3)
		} else {
			fmt.Println("Error:" + err.Error())
			fmt.Println("Press enter to quit.")
			bufio.NewReader(os.Stdin).ReadLine()
		}
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

	// int fmtOut
	if fmtEnable {
		initFormat(fmtOut)
	}

	// processing
	for _, pathSrc := range getFilelist(pathIn, &err) {
		if err != nil {
			return
		}
		if !strings.HasSuffix(pathSrc, ".xlsx") {
			continue
		} else if strings.HasPrefix(getBaseName(pathSrc), "~$") {
			continue
		}

		validFileCount++
		log("parsing file [%d][%s]\n", validFileCount, pathSrc)

		// Args to store tmp data
		vKey, vDesc, vData, vType := []string{}, []string{}, []string{}, []string{}

		// Analyze data from xlsx
		err = analyzeXlsx(pathSrc, &vKey, &vDesc, &vData, &vType)
		if err != nil {
			return
		}

		// Write data to xml
		if len(vKey) > 0 {
			pathTar := strings.TrimSuffix(pathOut+getRelativeDir(pathIn, pathSrc), ".xlsx") + ".xml"
			err = writeXml(pathTar, vKey, vDesc, vData, vType)
			if err != nil {
				return
			}

			if fmtEnable {
				pathTar = "." + strings.TrimSuffix(getRelativeDir(pathIn, pathSrc), ".xlsx") + ".xml"
				writeFormat(fmtOut, pathTar, vKey, vDesc, vType)
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
		fmt.Printf(format, args...)
	}
}

func loadConfig() (string, string, bool, string) {
	cfg := *conf.LoadConfig(config)
	format_file_enable, _ := strconv.ParseBool(cfg.Read("format_file", "enable"))
	return fixPath(cfg.Read("path", "in")), fixPath(cfg.Read("path", "out")), format_file_enable, fixPath(cfg.Read("format_file", "outFile"))
}

func fixPath(path string) string {
	return strings.Replace(path, "\\", "/", -1)
}

// get all files in the path
func getFilelist(path string, e *error) []string {
	fileV := []string{}
	filepath.Walk(path, func(path string, f os.FileInfo, err error) error {
		if f == nil {
			*e = err
			return nil
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

func getBaseName(path string) string {
	path = strings.Replace(path, "\\", "/", -1)
	if pos := strings.LastIndex(path, "/"); pos == -1 {
		return path
	} else {
		return path[pos+1:]
	}
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

func convertType(from string) string {
	if to, ok := Type_Conversion[from]; ok {
		return to
	}
	return from
}

func analyzeXlsx(pathSrc string, keyV, descV, dataV, typeV *[]string) error {
	// open file
	xlFile, err := xlsx.OpenFile(pathSrc)
	if err != nil {
		return err
	}

	// get valid columns
	mValid := make(map[int]bool)
	for _, sheet := range xlFile.Sheets {
		if len(sheet.Rows) < XData {
			return errors.New("No data in " + pathSrc)
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
	typeAllV := []string{}
	keyAllV := []string{}
	for _, sheet := range xlFile.Sheets {
		for x, row := range sheet.Rows {
			switch x {
			case XTitle:
			case XType:
				for y, cell := range row.Cells {
					text, _ := cell.String()
					typeAllV = append(typeAllV, text)
					if _, ok := mValid[y]; !ok {
						continue
					}
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
				dataValid := false
				dataS := "\t<data"
				for y, cell := range row.Cells {
					if _, ok := mValid[y]; !ok {
						continue
					} else if y >= len(keyAllV) {
						break
					}
					text, _ := cell.String()

					if fun, ok := Value_Conversion[typeAllV[y]]; ok {
						if !fun(text) {
							return errors.New(fmt.Sprintf("Invalid data [column:%s;row:%d;type:%s;data:%s]", (keyAllV)[y], x+1, typeAllV[y], text))
						}
					}
					dataS = fmt.Sprintf("%s %s=\"%s\"", dataS, (keyAllV)[y], text) // If slice can trans to array
					if len(text) > 0 {
						dataValid = true
					}
				}
				if dataValid {
					*dataV = append(*dataV, dataS+" />\n")
				}
			}
		}
		break
	}

	// check duplicate keys
	mKeyCount := make(map[string]int)
	for _, v := range keyAllV {
		mKeyCount[v]++
	}
	sDupKeys := ""
	for k, v := range mKeyCount {
		if v > 1 {
			sDupKeys += fmt.Sprintf("\t%s:%d;\n", k, v)
		}
	}
	if len(sDupKeys) > 0 {
		return errors.New("Duplicate key found:\n" + sDupKeys)
	} else {
		return nil
	}
}

func writeXml(pathTar string, keyV, descV, dataV, typeV []string) error {
	// check output floder
	dirname := getDirname(pathTar)
	if !isPathExist(dirname) {
		createPath(dirname)
	}

	// write file
	file, err := os.OpenFile(pathTar, os.O_CREATE|os.O_RDWR, 0664)
	if err != nil {
		return err
	}
	defer file.Close()
	if err := file.Truncate(0); err != nil {
		return err
	}
	writer := bufio.NewWriter(file)
	defer writer.Flush()

	// write head
	if _, err := writer.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"); err != nil {
		return err
	}

	// write comment
	if _, err := writer.WriteString("<!-- "); err != nil {
		return err
	}
	for k, v := range keyV {
		if k >= len(descV) {
			break
		}
		if _, err := writer.WriteString(fmt.Sprintf("%s=%s ", v, descV[k])); err != nil {
			return err
		}
	}
	if _, err := writer.WriteString("-->\n"); err != nil {
		return err
	}

	// write data
	if _, err := writer.WriteString("<root>\n"); err != nil {
		return err
	}
	for _, v := range dataV {
		_, err := writer.WriteString(v)
		if err != nil {
			return err
		}
	}
	if _, err := writer.WriteString("</root>\n"); err != nil {
		return err
	}
	return nil
}

func initFormat(pathTar string) {
	// check output floder
	dirname := getDirname(pathTar)
	if !isPathExist(dirname) {
		createPath(dirname)
	}

	// create/reset file
	file, err := os.OpenFile(pathTar, os.O_CREATE|os.O_RDWR, 0664)
	if err != nil {
		panic(err)
	}
	defer file.Close()
	if err := file.Truncate(0); err != nil {
		panic(err)
	}
}

func writeFormat(pathTar, fileName string, keyV, descV, typeV []string) {
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
	file, err := os.OpenFile(pathTar, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0664)
	if err != nil {
		panic(err)
	}
	defer file.Close()
	writer := bufio.NewWriter(file)
	defer writer.Flush()

	if _, err := writer.WriteString(fileName + "\n"); err != nil {
		panic(err)
	}
	for i := 0; i < length; i++ {
		if _, err := writer.WriteString(fmt.Sprintf("\t%s\t%s\t`xml:\"%s,attr\"`\t//%s\n", strings.ToUpper(keyV[i][0:1])+keyV[i][1:], convertType(typeV[i]), keyV[i], descV[i])); err != nil {
			panic(err)
		}
	}
	if _, err := writer.WriteString("\n"); err != nil {
		panic(err)
	}
}
