package xlsxp

import (
	"fmt"
	"testing"
)

func TestExportExcel(t *testing.T) {
	sheetName := "sheet1"
	type Test struct {
		CompletePercent float64 `json:"complete_percent" excel:"header(百分比);index(1);format(percent)"`
		UserName        string  `json:"user_name" excel:"header(学员);index(0);default(---)"`
	}

	datas := make([]Test, 0)
	datas = append(datas, Test{
		UserName:        "A",
		CompletePercent: 15.3,
	}, Test{
		UserName:        "B",
		CompletePercent: 17.558,
	})

	file, err := ExportExcel(sheetName, datas)
	if err != nil {
		t.Fatal(err)
	}

	for rIndex, row := range file.Sheet[sheetName].Rows {
		if rIndex == 0 {
			if row.Cells[0].String() != "学员" || row.Cells[1].String() != "百分比" {
				t.Fatal(fmt.Errorf("head not equal:%v", row.Cells))
			}
			continue
		}

		rowData := datas[rIndex-1]
		for cIndex, cell := range row.Cells {
			if cIndex == 0 {
				userName := rowData.UserName
				if userName != cell.String() {
					t.Fatal(fmt.Errorf("user_name not equal data:%s xlxs:%s",
						userName, cell.String()))
				}
			} else if cIndex == 1 {
				targetPercent := fmt.Sprintf("%.2f%%", rowData.CompletePercent)
				if targetPercent != cell.String() {
					t.Fatal(fmt.Errorf("complete_percent not equal data:%s xlxs:%s",
						targetPercent, cell.String()))
				}
			}
		}
	}
}
