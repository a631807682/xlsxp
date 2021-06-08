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
func ExportExcel(file *xlsx.File, sheetName string, vals interface{}, cformats ...CustomFormat) (err error) {
	// 初始化格式化函数
	formatFnMap := NewFormatFn(cformats...)

	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return
	}

	// 写入数据到excel
	err = setStructIntoExcelVals(sheet, vals, formatFnMap)
	if err != nil {
		return
	}

	return
}

type xCellField struct {
	Header       string   // 表头
	Width        float64  // 表头长度
	Index        int      // 排序
	DefaultValue string   // 导出默认值
	FormatFunc   FormatFn // 格式化函数
	FieldName    string   //字段名称
}

type xCellFields []xCellField

func (a xCellFields) Len() int           { return len(a) }
func (a xCellFields) Swap(i, j int)      { a[i], a[j] = a[j], a[i] }
func (a xCellFields) Less(i, j int) bool { return a[i].Index < a[j].Index }

func setStructIntoExcelVals(sheet *xlsx.Sheet, vals interface{}, formatFnMap FormatFnMap) (err error) {
	cellFieldMap, err := getStructFieldInfo(vals, formatFnMap)
	if err != nil {
		return
	}

	// 处理表头
	xCellFieldSlice := make(xCellFields, 0)
	for _, field := range cellFieldMap {
		xCellFieldSlice = append(xCellFieldSlice, field)
	}
	sort.Sort(xCellFieldSlice)

	rowVal := sheet.AddRow()
	for i, field := range xCellFieldSlice {
		excelCell := rowVal.AddCell()
		excelCell.SetValue(field.Header)
		if field.Width > 0 {
			sheet.SetColWidth(i, i, field.Width)
		}
	}

	// 处理表数据
	datasVal := reflect.ValueOf(vals)
	for rowIndex := 0; rowIndex < datasVal.Len(); rowIndex++ { // 行
		row := sheet.AddRow()
		rowVal := reflect.Indirect(datasVal.Index(rowIndex))

		for _, fieldInfo := range xCellFieldSlice {
			cell := row.AddCell()
			fieldValue := rowVal.FieldByName(fieldInfo.FieldName)
			if polyfillIsZero(fieldValue) && fieldInfo.DefaultValue != "" { //默认值
				fieldValue = reflect.ValueOf(fieldInfo.DefaultValue)
			}

			if fieldInfo.FormatFunc != nil {
				formatVal := fieldInfo.FormatFunc(fieldValue.Interface())
				fieldValue = reflect.ValueOf(formatVal)
			}

			cell.SetValue(fieldValue.Interface())
		}
	}

	return
}

// 获取结构体信息
func getStructFieldInfo(vals interface{}, formatFnMap FormatFnMap) (cellFieldMap map[int]xCellField, err error) {
	datasVal := reflect.ValueOf(vals)
	datasInd := reflect.Indirect(datasVal)
	datasKind := datasVal.Kind()
	if datasKind != reflect.Array && datasKind != reflect.Slice {
		err = fmt.Errorf("datas not array or slice")
		return
	}

	itemType := datasInd.Type().Elem()
	cellFieldMap = make(map[int]xCellField, 0)
	for fieldIndex := 0; fieldIndex < itemType.NumField(); fieldIndex++ {
		excelTag := itemType.Field(fieldIndex).Tag.Get(defaultStructTagName)
		if excelTag != "" { // 只处理定义了 `excel:"xxx"`
			_, tags := parseStructTag(excelTag)

			if headerIndex, ok := tags[tagIndex]; ok {
				headerName, _ := tags[tagHeader] // 表头
				defVal, _ := tags[tagDefault]    //默认值

				var cellWidth float64
				if widthStr, ok := tags[tagWidth]; ok { //表头宽度
					if val, err := strconv.ParseFloat(widthStr, 64); err == nil {
						cellWidth = val
					}
				}

				var formatFunc FormatFn
				if fnName, ok := tags[tagFormat]; ok { //格式化值
					formatFunc, ok = formatFnMap[fnName]
					if !ok {
						err = fmt.Errorf("excel tag format func not exist, check common format func or use defined custom func")
						return
					}
				}

				index, _ := strconv.Atoi(headerIndex)
				cellFiled := xCellField{
					Index:        index,
					Header:       headerName,
					Width:        cellWidth,
					DefaultValue: defVal,
					FormatFunc:   formatFunc,
					FieldName:    itemType.Field(fieldIndex).Name,
				}
				cellFieldMap[cellFiled.Index] = cellFiled
			}
		}
	}
	return
}
