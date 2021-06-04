package xlsxp

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/tealeg/xlsx"
)

func TestImportExcel(t *testing.T) {
	sheetName := "sheet1"
	type Test struct {
		CompletePercent float64 `json:"complete_percent" excel:"index(1);parse(percent)"`
		UserName        string  `json:"user_name" excel:"index(0);"`
	}

	file := xlsx.NewFile()
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		t.Fatal(err)
	}

	row := sheet.AddRow()
	row.AddCell().Value = "学员"
	row.AddCell().Value = "百分比"

	row = sheet.AddRow()
	row.AddCell().Value = "A"
	row.AddCell().Value = "15.3%"

	row = sheet.AddRow()
	row.AddCell().Value = "B"
	row.AddCell().Value = "17.55%"

	var xlsxBuf bytes.Buffer
	err = file.Write(&xlsxBuf)
	if err != nil {
		t.Fatal(err)
	}

	targetDatas := make([]Test, 0)
	err = ImportExcel(xlsxBuf.Bytes(), sheetName, &targetDatas)
	if err != nil {
		t.Fatal(err)
	}

	for i, data := range targetDatas {
		if i == 0 {
			if data.UserName != "A" {
				t.Fatal(fmt.Errorf("user_name not equal i:%d data:%s ", i, data.UserName))
			} else if data.CompletePercent != 15.3 {
				t.Fatal(fmt.Errorf("complete_percent not equal %d data:%f ", i, data.CompletePercent))
			}
		} else if i == 1 {
			if data.UserName != "B" {
				t.Fatal(fmt.Errorf("user_name not equal i:%d data:%s ", i, data.UserName))
			} else if data.CompletePercent != 17.55 {
				t.Fatal(fmt.Errorf("complete_percent not equal %d data:%f ", i, data.CompletePercent))
			}
		}
	}
}
