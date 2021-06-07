package xlsxp

import (
	"bytes"
	"fmt"
	"math"
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

// func TestExportExcel(t *testing.T) {
// 	sheetName := "sheet1"
// 	type Test struct {
// 		CompletePercent float64 `json:"complete_percent" excel:"header(百分比);index(1);format(percent)"`
// 		UserName        string  `json:"user_name" excel:"header(学员);index(0);default(---)"`
// 	}

// 	datas := make([]Test, 0)
// 	datas = append(datas, Test{
// 		UserName:        "A",
// 		CompletePercent: 15.3,
// 	}, Test{
// 		UserName:        "B",
// 		CompletePercent: 17.558,
// 	})

// 	file, err := ExportExcel(sheetName, datas)
// 	if err != nil {
// 		t.Fatal(err)
// 	}

// 	for rIndex, row := range file.Sheet[sheetName].Rows {
// 		if rIndex == 0 {
// 			if row.Cells[0].String() != "学员" || row.Cells[1].String() != "百分比" {
// 				t.Fatal(fmt.Errorf("head not equal:%v", row.Cells))
// 			}
// 			continue
// 		}

// 		rowData := datas[rIndex-1]
// 		for cIndex, cell := range row.Cells {
// 			if cIndex == 0 {
// 				userName := rowData.UserName
// 				if userName != cell.String() {
// 					t.Fatal(fmt.Errorf("user_name not equal data:%s xlxs:%s",
// 						userName, cell.String()))
// 				}
// 			} else if cIndex == 1 {
// 				targetPercent := fmt.Sprintf("%.2f%%", rowData.CompletePercent)
// 				if targetPercent != cell.String() {
// 					t.Fatal(fmt.Errorf("complete_percent not equal data:%s xlxs:%s",
// 						targetPercent, cell.String()))
// 				}
// 			}
// 		}
// 	}
// }
