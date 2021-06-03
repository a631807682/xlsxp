package xlsxp

import (
	"fmt"
	"reflect"
	"sort"
	"strconv"

	"github.com/tealeg/xlsx"
)

/* 导出 excel

sheetName: 表格名称

vals: 已定义 excel:"xx" 字段的结构数组数据

cformats: 自定义格式化函数

示例:

type Test struct {
	Test     string `json:"test" excel:"header(测试);index(1)"`
	UserName string `json:"user_name" excel:"header(学员);index(0)"`
}

| 学员         | 测试 			|
| ----------- | -----------    |
| 学员数据1    | 测试数据2       |
| 学员数据2    | 测试数据2       |

*/
func ExportExcel(sheetName string, vals interface{}, cformats ...CustomFormat) (file *xlsx.File, err error) {
	// 初始化格式化函数
	formatFnMap := NewFormatFn(cformats...)
	// 获取数据表头和值
	rows, err := getExcelKeyVals(vals, formatFnMap)
	if err != nil {
		return
	}

	file = xlsx.NewFile()
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return
	}

	rowVal := sheet.AddRow()
	for i, row := range rows {
		if i == 0 {
			for _, cell := range row {
				rowVal.AddCell().SetValue(cell.Header)
			}
		}

		rowVal = sheet.AddRow()
		for _, cell := range row {
			rowVal.AddCell().SetValue(cell.Cell)
		}
	}
	return
}

type xcell struct {
	Header string        // 表头
	Index  int           // 排序
	Cell   reflect.Value // 列数据
}

type xcells []xcell

func (a xcells) Len() int           { return len(a) }
func (a xcells) Swap(i, j int)      { a[i], a[j] = a[j], a[i] }
func (a xcells) Less(i, j int) bool { return a[i].Index < a[j].Index }

func getExcelKeyVals(vals interface{}, formatFnMap FormatFnMap) (rows [][]xcell, err error) {
	datas := reflect.ValueOf(vals)
	dataKind := datas.Kind()
	if dataKind != reflect.Array && dataKind != reflect.Slice {
		err = fmt.Errorf("datas not array or slice")
		return
	}

	rows = make([][]xcell, 0)
	for i := 0; i < datas.Len(); i++ { // 行
		rowVal := reflect.Indirect(datas.Index(i))
		rowType := rowVal.Type()

		cells := make([]xcell, 0)
		for j := 0; j < rowVal.NumField(); j++ { //列
			excelTag := rowType.Field(j).Tag.Get(defaultStructTagName)
			if excelTag != "" { // 只处理定义了 `excel:"xxx"`
				_, tags := parseStructTag(excelTag)
				if headerName, ok := tags[tagHeader]; ok { // 确定表头
					headerIndex, ok := tags[tagIndex]
					if !ok {
						err = fmt.Errorf("excel tag header miss index")
						return
					}

					cellVal := rowVal.Field(j)
					if defVal, ok := tags[tagDefault]; ok && polyfillIsZero(cellVal) { // 设置默认值
						cellVal = reflect.ValueOf(defVal)
					}

					if fnName, ok := tags[tagFormat]; ok { //格式化值
						fn, ok := formatFnMap[fnName]
						if !ok {
							err = fmt.Errorf("excel tag format func not exist, check common format func or use defined custom func")
							return
						}

						formatVal := fn(cellVal.Interface())
						cellVal = reflect.ValueOf(formatVal)
					}

					index, _ := strconv.Atoi(headerIndex)
					cell := xcell{
						Index:  index,
						Header: headerName,
						Cell:   cellVal,
					}
					cells = append(cells, cell)
				}
			}
		}
		sort.Sort(xcells(cells))
		rows = append(rows, cells)
	}
	return
}
