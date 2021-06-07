package xlsxp

import (
	"bytes"
	"fmt"
	"math"
	"math/rand"
	"reflect"
	"testing"
	"time"
)

type baseExportModel struct {
	Int   int   `excel:"index(0);header(Int)"`
	Int8  int8  `excel:"index(1);header(Int8)"`
	Int16 int16 `excel:"index(2);header(Int16)"`
	Int32 int32 `excel:"index(3);header(Int32)"`
	Int64 int64 `excel:"index(4);header(Int64)"`

	Uint   uint   `excel:"index(5);header(Uint)"`
	Uint8  uint8  `excel:"index(6);header(Uint8)"`
	Uint16 uint16 `excel:"index(7);header(Uint16)"`
	Uint32 uint32 `excel:"index(8);header(Uint32)"`
	Uint64 uint64 `excel:"index(9);header(Uint64)"`

	Float32 float32 `excel:"index(10);header(Float32)"`
	Float64 float64 `excel:"index(11);header(Float64)"`

	Byte   byte   `excel:"index(12);header(Byte)"`
	Rune   rune   `excel:"index(13);header(Rune)"`
	String string `excel:"index(14);header(String)"`
	Bool   bool   `excel:"index(15);header(Bool)"`
}

var baseExportData = baseExportModel{
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

func TestBaseTypeExportExcel(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]baseExportModel, 0)
	originData = append(originData, baseExportData, baseExportData)

	file, err := ExportExcel(sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	var xlsxBuf bytes.Buffer
	err = file.Write(&xlsxBuf)
	if err != nil {
		return
	}

	xlsxBytes := xlsxBuf.Bytes()

	targetDatas := make([]baseExportModel, 0)
	err = ImportExcel(xlsxBytes, sheetName, &targetDatas)
	if err != nil {
		t.Fatal(err)
	}

	if !reflect.DeepEqual(originData, targetDatas) {
		t.Fatal(fmt.Errorf("origin target not equal \n origin:%+v \n target:%+v",
			originData, targetDatas))
	}

}

type timeExportModel struct {
	Time time.Time `excel:"index(0)"` // 精度会丢失
}

func TestTimeTypeExportExcel(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]timeExportModel, 0)
	originData = append(originData, timeExportModel{
		Time: time.Now(),
	}, timeExportModel{
		Time: time.Now().Add(10000),
	})

	file, err := ExportExcel(sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	var xlsxBuf bytes.Buffer
	err = file.Write(&xlsxBuf)
	if err != nil {
		return
	}

	xlsxBytes := xlsxBuf.Bytes()
	targetDatas := make([]timeExportModel, 0)
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

const letterBytes = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

func randStringBytes(n int) string {
	b := make([]byte, n)
	for i := range b {
		b[i] = letterBytes[rand.Intn(len(letterBytes))]
	}
	return string(b)
}

type sortExportModel struct {
	StringA string `excel:"index(2)"`
	StringB string `excel:"index(1)"`
	StringC string `excel:"index(0)"`
}

func TestSortExportExcel(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]sortExportModel, 0)
	originData = append(originData, sortExportModel{
		StringA: randStringBytes(10),
		StringB: randStringBytes(10),
		StringC: randStringBytes(10),
	}, sortExportModel{
		StringA: randStringBytes(10),
		StringB: randStringBytes(10),
		StringC: randStringBytes(10),
	})

	file, err := ExportExcel(sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	var xlsxBuf bytes.Buffer
	err = file.Write(&xlsxBuf)
	if err != nil {
		return
	}

	xlsxBytes := xlsxBuf.Bytes()
	targetDatas := make([]sortExportModel, 0)
	err = ImportExcel(xlsxBytes, sheetName, &targetDatas)
	if err != nil {
		t.Fatal(err)
	}

	if !reflect.DeepEqual(originData, targetDatas) {
		t.Fatal(fmt.Errorf("origin target not equal \n origin:%+v \n target:%+v",
			originData, targetDatas))
	}
}

type formatExportOriginModel struct {
	Percent float64 `excel:"index(0);format(percent)"`
}

type formatExportTargetModel struct {
	PercentString string `excel:"index(0);"`
}

func TestFormatExportExcel(t *testing.T) {
	sheetName := "sheet1"

	originData := make([]formatExportOriginModel, 0)
	originData = append(originData, formatExportOriginModel{
		Percent: rand.Float64(),
	}, formatExportOriginModel{
		Percent: rand.Float64(),
	})

	file, err := ExportExcel(sheetName, originData)
	if err != nil {
		t.Fatal(err)
	}

	var xlsxBuf bytes.Buffer
	err = file.Write(&xlsxBuf)
	if err != nil {
		return
	}

	xlsxBytes := xlsxBuf.Bytes()
	targetDatas := make([]formatExportTargetModel, 0)
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

		targetPercent := fmt.Sprintf("%.2f%%", oData.Percent)
		if targetPercent != tData.PercentString {
			t.Fatal(fmt.Errorf("complete_percent not equal data:%s xlxs:%s",
				targetPercent, tData.PercentString))
		}
	}
}
