package xlsxp

import (
	"bytes"
	"fmt"
	"math"
	"reflect"
	"testing"
	"time"

	"github.com/tealeg/xlsx"
)

type baseImportModel struct {
	Int   int   `excel:"index(0)"`
	Int8  int8  `excel:"index(1)"`
	Int16 int16 `excel:"index(2)"`
	Int32 int32 `excel:"index(3)"`
	Int64 int64 `excel:"index(4)"`

	Uint   uint   `excel:"index(5)"`
	Uint8  uint8  `excel:"index(6)"`
	Uint16 uint16 `excel:"index(7)"`
	Uint32 uint32 `excel:"index(8)"`
	Uint64 uint64 `excel:"index(9)"`

	Float32 float32 `excel:"index(10)"`
	Float64 float64 `excel:"index(11)"`

	Byte   byte   `excel:"index(12)"`
	Rune   rune   `excel:"index(13)"`
	String string `excel:"index(14)"`
	Bool   bool   `excel:"index(15)"`
}

var baseImportData = baseImportModel{
	Int:    int(1<<31 - 1),
	Int8:   int8(1<<7 - 1),
	Int16:  int16(1<<15 - 1),
	Int32:  int32(1<<31 - 1),
	Int64:  int64(1<<63 - 1),
	Uint:   uint(1<<32 - 1),
	Uint8:  uint8(1<<8 - 1),
	Uint16: uint16(1<<16 - 1),
	Uint32: uint32(1<<32 - 1),
	Uint64: uint64(1<<63 - 1),

	Float32: float32(100.1234),
	Float64: float64(100.1234),

	Byte:   byte(1<<8 - 1),
	Rune:   rune(1<<31 - 1),
	String: "string",
	Bool:   true,
}

func genBaseXlsxBytes(datas interface{}, sheetName string) (xlsxDatas []byte, err error) {
	file := xlsx.NewFile()
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return
	}

	dv := reflect.ValueOf(datas)
	if dv.Kind() != reflect.Array && dv.Kind() != reflect.Slice {
		err = fmt.Errorf("datas not array or slice")
		return
	}

	sheet.AddRow() //表头
	for i := 0; i < dv.Len(); i++ {
		row := sheet.AddRow()
		v := dv.Index(i)
		for i := 0; i < v.NumField(); i++ {
			row.AddCell().SetValue(v.Field(i).Interface())
		}
	}

	var xlsxBuf bytes.Buffer
	err = file.Write(&xlsxBuf)
	if err != nil {
		return
	}
	xlsxDatas = xlsxBuf.Bytes()
	return
}

func TestBaseTypeImportExcel(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]baseImportModel, 0)
	originData = append(originData, baseImportData, baseImportData)
	xlsxBytes, err := genBaseXlsxBytes(originData, sheetName)
	if err != nil {
		t.Fatal(err)
	}

	targetDatas := make([]baseImportModel, 0)
	err = ImportExcel(xlsxBytes, sheetName, &targetDatas)
	if err != nil {
		t.Fatal(err)
	}

	if !reflect.DeepEqual(originData, targetDatas) {
		t.Fatal(fmt.Errorf("origin target not equal \n origin:%+v \n target:%+v",
			originData, targetDatas))
	}
}

type timeImportModel struct {
	Time time.Time `excel:"index(0)"` // 精度会丢失
}

func TestTimeTypeImportExcel(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]timeImportModel, 0)
	originData = append(originData, timeImportModel{
		Time: time.Now(),
	}, timeImportModel{
		Time: time.Now().Add(10000),
	})
	xlsxBytes, err := genBaseXlsxBytes(originData, sheetName)
	if err != nil {
		t.Fatal(err)
	}

	targetDatas := make([]timeImportModel, 0)
	err = ImportExcel(xlsxBytes, sheetName, &targetDatas)
	if err != nil {
		t.Fatal(err)
	}

	if len(originData) != len(targetDatas) {
		t.Fatal(fmt.Errorf("origin target len not equal"))
	}

	for i := 0; i < len(originData); i++ {
		oData := originData[i]
		tData := targetDatas[i]

		diffSeconds := oData.Time.Sub(tData.Time).Seconds()
		if math.Abs(diffSeconds) >= 1 {
			t.Fatal(fmt.Errorf("origin target time not equal"))
		}
	}
}
